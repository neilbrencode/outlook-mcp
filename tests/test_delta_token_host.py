"""The delta cursor is a URL the caller supplies and we fetch with a token.

``delta_token`` is a tool argument, so it comes from the model, and the model
reads email. Because ``_delta.fetch_delta_pages`` sends every URL with an
``Authorization: Bearer`` header, an unchecked cursor is a one-request leak of
a token scoped to the whole mailbox — not a mere request to the wrong place.

These are the guards for that. The important assertions are negative ones: on
a refused URL, *no request goes out* and *no token is minted*. A test that only
checked for a raised exception would still pass if the refusal happened after
the credential had already been handed to httpx.
"""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.errors import OutlookMCPError, ToolInputError
from outlook_mcp.tools._delta import fetch_delta_pages, require_graph_url

GRAPH_URL = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta"
GRAPH_DELTA_LINK = "https://graph.microsoft.com/v1.0/me/messages/delta?$deltatoken=abc"

# Each of these is a way to make a URL *look* like Graph to a human, a prefix
# test, or a substring test, while resolving somewhere else.
HOSTILE_URLS = [
    pytest.param("https://evil.example/steal", id="plain-other-host"),
    pytest.param("https://graph.microsoft.com@evil.example/steal", id="userinfo-prefix"),
    pytest.param("https://graph.microsoft.com.evil.example/steal", id="suffix-lookalike"),
    pytest.param("https://evil.example/graph.microsoft.com", id="host-in-path"),
    pytest.param("http://graph.microsoft.com/v1.0/me/messages/delta", id="plain-http"),
    pytest.param("http://169.254.169.254/latest/meta-data/", id="link-local-metadata"),
    pytest.param("file:///etc/passwd", id="file-scheme"),
    pytest.param("//evil.example/steal", id="protocol-relative"),
    pytest.param("https://GRAPH.MICROSOFT.COM.evil.example/x", id="case-lookalike"),
]


def _credential():
    """A credential that records whether anyone asked it for a token."""
    cred = MagicMock()
    cred.get_token = MagicMock(return_value=MagicMock(token="SECRET-TOKEN"))
    return cred


def _response(body: dict):
    r = MagicMock()
    r.status_code = 200
    r.json = MagicMock(return_value=body)
    r.raise_for_status = MagicMock()
    return r


def _patched_client(responses):
    """Patch httpx.AsyncClient, returning (patcher, client) so calls can be read."""
    queue = list(responses)

    async def fake_get(url, headers=None):
        return queue.pop(0)

    client = MagicMock()
    client.get = AsyncMock(side_effect=fake_get)
    client.__aenter__ = AsyncMock(return_value=client)
    client.__aexit__ = AsyncMock(return_value=False)
    return patch("outlook_mcp.tools._delta.httpx.AsyncClient", return_value=client), client


# ── require_graph_url ────────────────────────────────────────────────


@pytest.mark.parametrize("url", HOSTILE_URLS)
def test_require_graph_url_refuses_non_graph_hosts(url):
    with pytest.raises(ToolInputError):
        require_graph_url(url, source="delta_token")


@pytest.mark.parametrize(
    "url",
    [
        GRAPH_URL,
        GRAPH_DELTA_LINK,
        "https://graph.microsoft.com/v1.0/me/contacts/delta",
        # Host comparison is case-insensitive on the host itself.
        "https://GRAPH.microsoft.COM/v1.0/me/contacts/delta",
    ],
)
def test_require_graph_url_allows_graph(url):
    assert require_graph_url(url, source="delta_token") == url


def test_refusal_names_the_argument_and_reaches_the_model():
    """The text is what the agent reads, so it must say how to recover.

    ``ToolInputError`` inherits ``OutlookMCPError`` (and so the SDK's
    ``ToolError``); anything inheriting plain ``Exception`` would reach the
    client as a bare "Error executing tool" with the message withheld.
    """
    with pytest.raises(ToolInputError) as exc:
        require_graph_url("https://evil.example/x", source="delta_token")

    assert isinstance(exc.value, OutlookMCPError)
    assert isinstance(exc.value, ValueError)
    text = str(exc.value)
    assert "delta_token" in text
    assert "graph.microsoft.com" in text
    assert "evil.example" in text


# ── fetch_delta_pages: the refusal must precede the token ────────────


@pytest.mark.parametrize("url", HOSTILE_URLS)
async def test_hostile_delta_token_sends_nothing_and_mints_nothing(url):
    cred = _credential()
    patcher, client = _patched_client([])

    with patcher:
        with pytest.raises(ToolInputError):
            await fetch_delta_pages(cred, initial_url=GRAPH_URL, delta_token=url, page_size=10)

    client.get.assert_not_called()
    cred.get_token.assert_not_called()


async def test_hostile_nextlink_is_not_followed():
    """Graph's own response is not a trusted source of the next URL.

    If the first hop can be attacker-chosen, so can the ``@odata.nextLink``
    its response carries — and a compromised first hop that returns a nextLink
    would otherwise get a second request, with the same token, for free.
    """
    cred = _credential()
    first = _response(
        {
            "value": [{"id": "1"}],
            "@odata.nextLink": "https://evil.example/page2",
        }
    )
    second = _response({"value": [], "@odata.deltaLink": GRAPH_DELTA_LINK})
    patcher, client = _patched_client([first, second])

    with patcher:
        with pytest.raises(ToolInputError):
            await fetch_delta_pages(cred, initial_url=GRAPH_URL, delta_token=None, page_size=10)

    # The first (legitimate) request went out; the poisoned second did not.
    assert client.get.await_count == 1


async def test_hostile_deltalink_is_not_handed_back_as_a_token():
    """A poisoned deltaLink must fail now, not on the caller's next call."""
    cred = _credential()
    patcher, client = _patched_client(
        [_response({"value": [], "@odata.deltaLink": "https://evil.example/d"})]
    )

    with patcher:
        with pytest.raises(ToolInputError):
            await fetch_delta_pages(cred, initial_url=GRAPH_URL, delta_token=None, page_size=10)


# ── the happy path still works ───────────────────────────────────────


async def test_graph_delta_token_is_used_verbatim():
    cred = _credential()
    patcher, client = _patched_client(
        [_response({"value": [{"id": "1"}], "@odata.deltaLink": GRAPH_DELTA_LINK})]
    )

    with patcher:
        items, token, has_more = await fetch_delta_pages(
            cred, initial_url="", delta_token=GRAPH_DELTA_LINK, page_size=10
        )

    assert items == [{"id": "1"}]
    assert token == GRAPH_DELTA_LINK
    assert has_more is False
    assert client.get.await_args.args[0] == GRAPH_DELTA_LINK
    assert client.get.await_args.kwargs["headers"]["Authorization"] == ("Bearer SECRET-TOKEN")


async def test_graph_nextlink_is_still_followed():
    cred = _credential()
    next_link = "https://graph.microsoft.com/v1.0/me/messages/delta?$skiptoken=xyz"
    patcher, client = _patched_client(
        [
            _response({"value": [{"id": "1"}], "@odata.nextLink": next_link}),
            _response({"value": [{"id": "2"}], "@odata.deltaLink": GRAPH_DELTA_LINK}),
        ]
    )

    with patcher:
        items, token, has_more = await fetch_delta_pages(
            cred, initial_url=GRAPH_URL, delta_token=None, page_size=10
        )

    assert [i["id"] for i in items] == ["1", "2"]
    assert token == GRAPH_DELTA_LINK
    assert client.get.await_count == 2


# ── every tool that exposes the argument ─────────────────────────────


async def test_list_inbox_delta_refuses_hostile_token():
    from outlook_mcp.tools.mail_delta import list_inbox_delta

    graph = MagicMock()
    graph.credential = _credential()
    patcher, client = _patched_client([])

    with patcher:
        with pytest.raises(ToolInputError):
            await list_inbox_delta(graph, delta_token="https://evil.example/steal")

    client.get.assert_not_called()
    graph.credential.get_token.assert_not_called()


async def test_list_events_delta_refuses_hostile_token():
    from outlook_mcp.tools.calendar_delta import list_events_delta

    graph = MagicMock()
    graph.credential = _credential()
    patcher, client = _patched_client([])

    with patcher:
        with pytest.raises(ToolInputError):
            await list_events_delta(graph, None, None, delta_token="https://evil.example/steal")

    client.get.assert_not_called()
    graph.credential.get_token.assert_not_called()


async def test_list_contacts_delta_refuses_hostile_token():
    from outlook_mcp.tools.contacts_delta import list_contacts_delta

    graph = MagicMock()
    graph.credential = _credential()
    patcher, client = _patched_client([])

    with patcher:
        with pytest.raises(ToolInputError):
            await list_contacts_delta(graph, delta_token="https://evil.example/steal")

    client.get.assert_not_called()
    graph.credential.get_token.assert_not_called()


@pytest.mark.parametrize("resource", ["mail", "events", "contacts"])
async def test_changes_since_refuses_hostile_token(resource):
    """The digest takes a dict of tokens; each value is the same sink.

    Its three resources run concurrently, so the two holding no token still
    make their legitimate bootstrap calls. The assertion is therefore not
    "nothing was sent" but "nothing was sent *there*" — and the refusal must
    still propagate out of the gather rather than being absorbed as a stale
    token, which would turn a blocked exfiltration attempt into a call that
    looks like it worked.
    """
    from outlook_mcp.tools.digest import changes_since

    graph = MagicMock()
    graph.credential = _credential()
    requested: list[str] = []

    async def fake_get(url, headers=None):
        requested.append(url)
        return _response({"value": [], "@odata.deltaLink": GRAPH_DELTA_LINK})

    client = MagicMock()
    client.get = AsyncMock(side_effect=fake_get)
    client.__aenter__ = AsyncMock(return_value=client)
    client.__aexit__ = AsyncMock(return_value=False)

    with patch("outlook_mcp.tools._delta.httpx.AsyncClient", return_value=client):
        with pytest.raises(ToolInputError):
            await changes_since(graph, delta_tokens={resource: "https://evil.example/steal"})

    assert not any("evil.example" in url for url in requested), requested
