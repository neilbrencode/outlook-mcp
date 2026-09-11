"""Shared helpers for Graph ``$delta`` queries.

These tools deliberately bypass the msgraph SDK and hit Graph's delta
endpoints with raw httpx. The SDK's typed delta builders rebuild URL
templates from query-parameter dataclasses (so a deltaLink passed through
them gets its ``$deltatoken`` stripped) and discard the ``@removed``
annotation on returned items, which delta-query callers need to detect
tombstones. Raw httpx is also easier to wire to the per-endpoint quirks
documented inline:

- ``/me/calendarView/delta`` requires ``startDateTime`` + ``endDateTime``
  on the first call but rejects ``$top``; you must pass
  ``Prefer: odata.maxpagesize=N`` instead.
- ``/me/contacts/delta`` accepts no query string at all on the first call;
  ``Prefer: odata.maxpagesize=N`` is again how you size the page.
- ``/me/mailFolders/{id}/messages/delta`` accepts ``$top``.

Tokens are *not* persisted by outlook-mcp — that's the caller's job (an
agent decides where it wants to store its watermark). We pass the raw
``@odata.deltaLink`` / ``@odata.nextLink`` URL through as the opaque
cursor string, which means the cursor an agent hands back is a URL this
module then fetches *with a bearer token attached*. Every URL fetched here
is therefore host-checked first — see ``require_graph_url``.
"""

from __future__ import annotations

from typing import Any
from urllib.parse import urlparse

import httpx

from outlook_mcp.errors import ToolInputError
from outlook_mcp.throttle import send_with_retry

GRAPH_BASE = "https://graph.microsoft.com/v1.0/"
GRAPH_TOKEN_SCOPE = "https://graph.microsoft.com/.default"

# The only host this module will send a bearer token to. Exact match, not a
# suffix test: `graph.microsoft.com.evil.example` ends with the right string.
GRAPH_HOST = "graph.microsoft.com"

# Safety cap multiplier — bound a single tool call to at most this many
# items even when Graph keeps handing us more ``@odata.nextLink`` pages
# inside one delta-sync round. Callers continue by passing the returned
# delta_token (a nextLink) back in.
PAGE_SIZE_CAP_MULTIPLIER = 4


def require_graph_url(url: str, *, source: str) -> str:
    """Return ``url`` if it is an https Microsoft Graph URL, else refuse.

    ``delta_token`` is a tool argument, so it arrives from the model — and the
    model reads email. "Your sync state was reset, resume from <url>" in a
    message body is a working instruction, and every URL this module fetches
    carries the caller's Graph bearer token in an ``Authorization`` header. An
    unchecked host is therefore not merely a request to somewhere unexpected:
    it hands a token scoped to the whole mailbox (``.default`` consents every
    scope on the app registration, so ``read_only`` does not shrink it) to
    whoever chose the URL. This is the network-side twin of the confinement
    ``resolve_attachment_path`` already applies to filesystem paths.

    ``@odata.nextLink`` gets the same treatment. Once the first hop can be
    attacker-chosen, so can the link its response carries.

    Matching is on the parsed host, never on a prefix or a substring:
    ``https://graph.microsoft.com@evil.example/x`` parses to host
    ``evil.example``, and ``graph.microsoft.com.evil.example`` is not
    ``graph.microsoft.com``.

    Only the global-cloud host is accepted. This server targets personal
    accounts (tenant ``consumers``), which exist nowhere else, so the
    sovereign-cloud hostnames would be surface with no caller.
    """
    try:
        parsed = urlparse(url)
    except ValueError as e:
        raise ToolInputError(f"{source} is not a valid URL: {url[:100]!r}") from e

    if parsed.scheme != "https" or parsed.hostname != GRAPH_HOST:
        raise ToolInputError(
            f"{source} must be an https://{GRAPH_HOST} URL produced by a prior "
            f"delta call; got {url[:100]!r}. Pass back the delta_token from a "
            "previous response verbatim, or omit it to start a new sync round."
        )
    return url


def _bearer_token(credential: Any) -> str:
    """Mint a Graph access token from an azure-identity credential."""
    tok = credential.get_token(GRAPH_TOKEN_SCOPE)
    return tok.token


def format_delta_item(raw: dict, normal_formatter) -> dict:
    """Map a raw Graph delta-response item to outlook-mcp's wire shape.

    Items annotated with ``@removed`` are tombstones — Graph returns only
    the ``id`` (and the annotation), no other fields. We collapse those
    to ``{id, is_deleted: True}`` so callers don't have to special-case
    sparse rows.

    Live items go through ``normal_formatter`` (which expects the dict
    shape Graph sends — *not* an SDK object) and get ``is_deleted: False``
    appended.
    """
    if "@removed" in raw:
        return {"id": raw.get("id"), "is_deleted": True}
    out = normal_formatter(raw)
    out["is_deleted"] = False
    return out


async def fetch_delta_pages(
    credential: Any,
    *,
    initial_url: str,
    delta_token: str | None,
    page_size: int,
    headers: dict[str, str] | None = None,
    timeout: float = 30.0,
) -> tuple[list[dict], str | None, bool]:
    """Walk Graph delta pages until a ``deltaLink`` or the safety cap.

    Behavior:

    - When ``delta_token`` is provided it's used as the request URL
      verbatim — it's already a full Graph URL (either a deltaLink that
      starts a new sync round or a nextLink that resumes an in-progress
      one). ``initial_url`` is ignored in that case.
    - When ``delta_token`` is ``None`` we hit ``initial_url`` (the
      caller-built first-sync URL with whatever query params and
      Prefer headers the resource accepts).
    - We auto-follow ``@odata.nextLink`` inside a single tool call up to
      ``page_size * PAGE_SIZE_CAP_MULTIPLIER`` items so an agent isn't
      forced to chain four follow-up calls just to drain Graph's
      pagination on a large initial snapshot.
    - We stop early on the cap and surface ``has_more=True`` plus the
      nextLink as the returned token so the caller can resume.
    - We stop on a ``deltaLink`` and surface ``has_more=False`` plus the
      deltaLink as the returned token — the caller stores that and uses
      it for the *next* sync round.
    - If Graph returns neither (rare but valid: zero changes plus no
      deltaLink, which means the previous token is still valid), we
      return ``has_more=False`` and ``delta_token=None`` so the caller
      keeps using their old token.

    Returns ``(raw_items, next_token, has_more)``. ``raw_items`` is the
    accumulated list of dicts straight from Graph; the per-resource
    wrapper is responsible for mapping them through ``format_delta_item``
    with the right per-item formatter.
    """
    if page_size < 1:
        page_size = 1
    cap = page_size * PAGE_SIZE_CAP_MULTIPLIER

    # Before the token is minted, not after: a refused URL must never cause a
    # bearer token to exist on this code path at all.
    url: str = (
        require_graph_url(delta_token, source="delta_token")
        if delta_token
        else require_graph_url(initial_url, source="initial_url")
    )
    base_headers = {
        "Authorization": f"Bearer {_bearer_token(credential)}",
        "Accept": "application/json",
        # httpx decompresses transparently; delta pages are large and highly
        # compressible, and polling agents fetch them constantly.
        "Accept-Encoding": "gzip",
    }
    if headers:
        base_headers.update(headers)

    collected: list[dict] = []
    next_token: str | None = None
    has_more = False

    async with httpx.AsyncClient(timeout=timeout) as client:
        while True:
            # Retry the delta GET on 429/503 (this raw-httpx path bypasses the
            # SDK's kiota RetryHandler, so it must honor Retry-After itself).
            r = await send_with_retry(client, "GET", url, headers=base_headers)
            r.raise_for_status()
            body = r.json()

            page_items = body.get("value") or []
            collected.extend(page_items)

            next_link = body.get("@odata.nextLink")
            delta_link = body.get("@odata.deltaLink")

            # Checked on the way in, so a poisoned link fails here rather than
            # being handed back to the caller as a token that only explodes on
            # some later call.
            if next_link:
                next_link = require_graph_url(next_link, source="@odata.nextLink")
            if delta_link:
                delta_link = require_graph_url(delta_link, source="@odata.deltaLink")

            if delta_link:
                # Reached the end of this sync round. The deltaLink is the
                # cursor for the *next* round.
                next_token = delta_link
                has_more = False
                break

            if next_link:
                if len(collected) >= cap:
                    # Hit the per-call cap mid-sync. Hand the nextLink back
                    # so the caller resumes from where we stopped.
                    next_token = next_link
                    has_more = True
                    break
                url = next_link
                continue

            # Neither link present — possible on the no-changes path when
            # Graph echoes the old window without a new deltaLink. Caller
            # should reuse their previous token.
            next_token = None
            has_more = False
            break

    return collected, next_token, has_more
