#!/usr/bin/env python3
"""
Microsoft To Do – Complete MCP Server
======================================
A Model Context Protocol (MCP) server that gives any MCP-compatible AI client
full read/write access to Microsoft To Do via Microsoft Graph API.

30 tools covering every To Do endpoint:
  • Task Lists          – list, create, get, update, delete
  • Tasks               – list, create, get, update, delete, complete
  • Checklist Items     – list, create, get, update, delete
  • Linked Resources    – list, create, get, update, delete
  • Attachments (beta)  – list, create, get, delete
  • Convenience         – find_list_by_name, find_task_by_title
  • Sync                – delta query for incremental task sync
  • Auth                – start_auth, finish_auth (first-time setup)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 SETUP (one-time, ~5 minutes)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Step 1 – Register a free app (no credit card, no paid plan):

  1. Go to https://entra.microsoft.com and sign in with your Microsoft account
     (the same account you use for To Do)
  2. Identity → Applications → App registrations → New registration
  3. Name: "MS To Do MCP"
     Supported account types: "Personal Microsoft accounts only"
  4. Click Register → copy the Application (client) ID
  5. Authentication → Add a platform → Mobile and desktop applications
     → check: https://login.microsoftonline.com/common/oauth2/nativeclient → Configure
     → Advanced settings: "Allow public client flows" → Yes → Save
  6. API permissions → Add a permission → Microsoft Graph → Delegated
     → add Tasks.ReadWrite and User.Read → Grant admin consent

Step 2 – Get your refresh token (run locally):

    pip install "mcp[cli]" httpx
    export AZURE_CLIENT_ID=<paste-client-id-from-step-1>
    python auth.py
    # Follow the sign-in instructions printed to the terminal
    # Copy the GRAPH_REFRESH_TOKEN printed at the end

Step 3 – Deploy to Fly.io:

    fly launch --copy-config --yes
    fly secrets set AZURE_CLIENT_ID=<...> GRAPH_REFRESH_TOKEN=<...>

    Your MCP URL: https://<app-name>.fly.dev/mcp

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 LOCAL DEVELOPMENT
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    export AZURE_CLIENT_ID=<...>
    export GRAPH_REFRESH_TOKEN=<...>
    python microsoft_todo_mcp_server.py
    # Starts on http://localhost:3000/mcp

Claude Desktop (add to claude_desktop_config.json):
    {
      "mcpServers": {
        "microsoft-todo": {
          "command": "python",
          "args": ["/full/path/to/microsoft_todo_mcp_server.py"],
          "env": {
            "AZURE_CLIENT_ID": "...",
            "GRAPH_REFRESH_TOKEN": "...",
            "MCP_TRANSPORT": "stdio"
          }
        }
      }
    }

Local Development
-----------------
    pip install "mcp[cli]" httpx
    export GRAPH_ACCESS_TOKEN="eyJ0eX..."
    python microsoft_todo_mcp_server.py

    # Server starts on http://localhost:3000/mcp

Claude Desktop – add to claude_desktop_config.json:
    {
      "mcpServers": {
        "microsoft-todo": {
          "command": "python",
          "args": ["/full/path/to/microsoft_todo_mcp_server.py"],
          "env": {
            "AZURE_CLIENT_ID": "...",
            "AZURE_TENANT_ID": "...",
            "AZURE_CLIENT_SECRET": "...",
            "MCP_TRANSPORT": "stdio"
          }
        }
      }
    }

Cursor / VS Code – add to .cursor/mcp.json:
    {
      "mcpServers": {
        "microsoft-todo": {
          "url": "http://localhost:3000/mcp"
        }
      }
    }
"""

from __future__ import annotations

import json
import os
import sys
import time
import logging
from typing import Any, Optional

import httpx
from mcp.server.fastmcp import FastMCP

# ─── Logging (stderr only – stdout is reserved for MCP JSON-RPC) ───
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("todo-mcp")

# ─── Constants ───
GRAPH_V1 = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"
SCOPE = "Tasks.ReadWrite User.Read offline_access"
TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
DEVICE_CODE_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode"

# ─── MCP Server ───
mcp = FastMCP(
    "Microsoft To Do",
    description=(
        "Full read/write access to Microsoft To Do. "
        "Manage task lists, tasks, checklist items (steps), "
        "linked resources, and attachments."
    ),
)


# ═══════════════════════════════════════════════════════════
#  AUTH HELPERS
# ═══════════════════════════════════════════════════════════

_token_cache: dict[str, Any] = {}


async def _get_token() -> str:
    """Return a valid delegated access token.

    Priority:
      1. GRAPH_ACCESS_TOKEN env var (static, for quick testing — expires ~1h)
      2. GRAPH_REFRESH_TOKEN + AZURE_CLIENT_ID (production — auto-refreshes)

    Run auth.py once locally to obtain GRAPH_REFRESH_TOKEN.
    """
    static = os.environ.get("GRAPH_ACCESS_TOKEN")
    if static:
        return static

    client_id = os.environ.get("AZURE_CLIENT_ID", "")
    refresh_token = os.environ.get("GRAPH_REFRESH_TOKEN", "")

    if not client_id or not refresh_token:
        raise RuntimeError(
            "Authentication not configured. "
            "Run auth.py locally to get GRAPH_REFRESH_TOKEN, then set:\n"
            "  AZURE_CLIENT_ID=<your-app-client-id>\n"
            "  GRAPH_REFRESH_TOKEN=<token-from-auth.py>"
        )

    if _token_cache.get("expires_at", 0) > time.time() + 60:
        return _token_cache["access_token"]

    async with httpx.AsyncClient() as client:
        resp = await client.post(TOKEN_URL, data={
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": SCOPE,
        })
        resp.raise_for_status()
        body = resp.json()

    _token_cache["access_token"] = body["access_token"]
    _token_cache["expires_at"] = time.time() + body.get("expires_in", 3600)
    if "refresh_token" in body:
        _token_cache["refresh_token"] = body["refresh_token"]
        log.info("Refresh token rotated — update GRAPH_REFRESH_TOKEN env var if needed")

    return _token_cache["access_token"]


# ═══════════════════════════════════════════════════════════
#  SETUP TOOL (device code flow for first-time auth)
# ═══════════════════════════════════════════════════════════

_pending_device: dict[str, Any] = {}


@mcp.tool()
async def start_auth(client_id: str) -> str:
    """Start the one-time Microsoft sign-in flow (device code).

    Call this the very first time to get a sign-in URL and code.
    Then open the URL in your browser, enter the code, and sign in
    with your personal Microsoft account. Once done, call finish_auth
    to get your GRAPH_REFRESH_TOKEN to save as an env var.

    Args:
        client_id: Your Azure app's Client ID (from entra.microsoft.com).
    """
    async with httpx.AsyncClient() as client:
        resp = await client.post(DEVICE_CODE_URL, data={
            "client_id": client_id,
            "scope": SCOPE,
        })
        resp.raise_for_status()
        data = resp.json()

    _pending_device.update(data)
    _pending_device["client_id"] = client_id

    return json.dumps({
        "instructions": [
            f"1. Open this URL: {data['verification_uri']}",
            f"2. Enter this code: {data['user_code']}",
            "3. Sign in with your Microsoft account (the one you use for To Do)",
            "4. Come back and call finish_auth to get your refresh token",
        ],
        "url": data["verification_uri"],
        "code": data["user_code"],
        "expires_in_seconds": data.get("expires_in", 900),
    }, indent=2)


@mcp.tool()
async def finish_auth() -> str:
    """Complete the sign-in started by start_auth and return your refresh token.

    Call this after you have signed in at the URL from start_auth.
    Save the GRAPH_REFRESH_TOKEN value as a permanent env var.
    """
    if not _pending_device:
        return json.dumps({"error": "Call start_auth first."})

    client_id = _pending_device["client_id"]
    device_code = _pending_device["device_code"]
    deadline = time.time() + _pending_device.get("expires_in", 900)

    async with httpx.AsyncClient() as client:
        while time.time() < deadline:
            resp = await client.post(TOKEN_URL, data={
                "client_id": client_id,
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "device_code": device_code,
            })
            body = resp.json()

            if "access_token" in body:
                _pending_device.clear()
                return json.dumps({
                    "success": True,
                    "message": "Authenticated! Save these as env vars on your hosting platform:",
                    "AZURE_CLIENT_ID": client_id,
                    "GRAPH_REFRESH_TOKEN": body["refresh_token"],
                    "fly_command": (
                        f"fly secrets set "
                        f"AZURE_CLIENT_ID={client_id} "
                        f"GRAPH_REFRESH_TOKEN={body['refresh_token']}"
                    ),
                }, indent=2)

            error = body.get("error", "")
            if error == "authorization_pending":
                await __import__("asyncio").sleep(5)
            elif error == "slow_down":
                await __import__("asyncio").sleep(10)
            else:
                return json.dumps({"error": error, "detail": body})

    return json.dumps({"error": "Code expired. Call start_auth again."})


# ═══════════════════════════════════════════════════════════
#  HTTP HELPERS
# ═══════════════════════════════════════════════════════════

async def _graph(
    method: str,
    path: str,
    *,
    body: dict | None = None,
    params: dict | None = None,
    base: str = GRAPH_V1,
) -> dict | str:
    """Make an authenticated request to Microsoft Graph."""
    token = await _get_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    url = f"{base}{path}"
    log.info("%s %s", method, url)

    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.request(
            method, url, headers=headers, json=body, params=params
        )

    if resp.status_code == 204:
        return {"status": "success", "message": "Deleted successfully."}

    if resp.status_code >= 400:
        try:
            err = resp.json()
        except Exception:
            err = resp.text
        return {"error": True, "status_code": resp.status_code, "detail": err}

    return resp.json()


async def _graph_paged(
    path: str,
    *,
    params: dict | None = None,
    base: str = GRAPH_V1,
    max_pages: int = 20,
) -> dict:
    """Fetch all pages of a Graph list response, following @odata.nextLink."""
    token = await _get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    url: str | None = f"{base}{path}"
    all_values: list[Any] = []
    page = 0

    async with httpx.AsyncClient(timeout=30) as client:
        while url and page < max_pages:
            log.info("GET %s (page %d)", url, page + 1)
            resp = await client.get(url, headers=headers, params=params if page == 0 else None)
            if resp.status_code >= 400:
                try:
                    err = resp.json()
                except Exception:
                    err = resp.text
                return {"error": True, "status_code": resp.status_code, "detail": err}
            data = resp.json()
            all_values.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            page += 1

    return {"value": all_values, "@odata.count": len(all_values)}


def _build_datetime(date: str | None, tz: str | None) -> dict | None:
    if not date:
        return None
    return {"dateTime": date, "timeZone": tz or "UTC"}


# ═══════════════════════════════════════════════════════════
#  1. TASK LISTS
# ═══════════════════════════════════════════════════════════

@mcp.tool()
async def list_task_lists() -> str:
    """List all To Do task lists for the user.

    Returns every list including the default 'Tasks' list,
    flagged emails, and any custom lists.
    """
    result = await _graph_paged("/me/todo/lists")
    return json.dumps(result, indent=2)


@mcp.tool()
async def create_task_list(display_name: str) -> str:
    """Create a new To Do task list.

    Args:
        display_name: The name for the new list (e.g. "Next Actions", "Someday/Maybe").
    """
    result = await _graph("POST", "/me/todo/lists", body={"displayName": display_name})
    return json.dumps(result, indent=2)


@mcp.tool()
async def get_task_list(list_id: str) -> str:
    """Get a specific task list by its ID.

    Args:
        list_id: The unique identifier of the task list.
    """
    result = await _graph("GET", f"/me/todo/lists/{list_id}")
    return json.dumps(result, indent=2)


@mcp.tool()
async def update_task_list(list_id: str, display_name: str) -> str:
    """Rename a task list.

    Args:
        list_id: The unique identifier of the task list.
        display_name: The new name for the list.
    """
    result = await _graph("PATCH", f"/me/todo/lists/{list_id}", body={"displayName": display_name})
    return json.dumps(result, indent=2)


@mcp.tool()
async def delete_task_list(list_id: str) -> str:
    """Delete a task list and all of its tasks permanently.

    Args:
        list_id: The unique identifier of the task list to delete.
    """
    result = await _graph("DELETE", f"/me/todo/lists/{list_id}")
    return json.dumps(result, indent=2)


# ═══════════════════════════════════════════════════════════
#  2. TASKS
# ═══════════════════════════════════════════════════════════

@mcp.tool()
async def list_tasks(
    list_id: str,
    filter: Optional[str] = None,
    top: Optional[int] = None,
    orderby: Optional[str] = None,
    select: Optional[str] = None,
) -> str:
    """List all tasks in a task list. Supports OData query parameters.
    Automatically follows pagination to return all results.

    Args:
        list_id: The ID of the task list.
        filter: OData $filter expression, e.g. "status eq 'notStarted'".
        top: Maximum number of tasks to return (per page; all pages are fetched).
        orderby: OData $orderby expression, e.g. "createdDateTime desc".
        select: Comma-separated properties to include, e.g. "title,status,dueDateTime".
    """
    params: dict[str, Any] = {}
    if filter:
        params["$filter"] = filter
    if top:
        params["$top"] = top
    if orderby:
        params["$orderby"] = orderby
    if select:
        params["$select"] = select

    result = await _graph_paged(
        f"/me/todo/lists/{list_id}/tasks",
        params=params or None,
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def create_task(
    list_id: str,
    title: str,
    body_content: Optional[str] = None,
    body_content_type: Optional[str] = "text",
    importance: Optional[str] = "normal",
    status: Optional[str] = "notStarted",
    due_date: Optional[str] = None,
    due_timezone: Optional[str] = "UTC",
    reminder_date: Optional[str] = None,
    reminder_timezone: Optional[str] = "UTC",
    is_reminder_on: Optional[bool] = None,
    categories: Optional[list[str]] = None,
    start_date: Optional[str] = None,
    start_timezone: Optional[str] = "UTC",
    recurrence_pattern: Optional[str] = None,
    recurrence_interval: Optional[int] = None,
    recurrence_range_type: Optional[str] = None,
    recurrence_range_end_date: Optional[str] = None,
    recurrence_occurrences: Optional[int] = None,
) -> str:
    """Create a new task in a task list.

    Args:
        list_id: The ID of the task list to add the task to.
        title: The task title / description.
        body_content: Detailed notes for the task body (optional).
        body_content_type: Body format – "text" or "html" (default "text").
        importance: Priority – "low", "normal", or "high" (default "normal").
        status: Task status – "notStarted", "inProgress", "completed",
                "waitingOnOthers", or "deferred" (default "notStarted").
        due_date: Due date in ISO 8601 format, e.g. "2026-05-01T09:00:00".
        due_timezone: Time zone for due date, e.g. "AUS Eastern Standard Time".
        reminder_date: Reminder date/time in ISO 8601 format.
        reminder_timezone: Time zone for the reminder.
        is_reminder_on: Whether to enable the reminder (set True if reminder_date provided).
        categories: List of category labels, e.g. ["Important", "Work"].
        start_date: Start date in ISO 8601 format (optional).
        start_timezone: Time zone for start date.
        recurrence_pattern: Recurrence type – "daily", "weekly", "absoluteMonthly",
                            "absoluteYearly", "relativeMonthly", or "relativeYearly".
        recurrence_interval: How often the pattern repeats (e.g. 1 = every week).
        recurrence_range_type: "endDate", "noEnd", or "numbered".
        recurrence_range_end_date: End date for recurrence (ISO 8601 date, e.g. "2026-12-31").
        recurrence_occurrences: Number of occurrences (used when recurrence_range_type="numbered").
    """
    task: dict[str, Any] = {"title": title}

    if body_content:
        task["body"] = {"content": body_content, "contentType": body_content_type}
    if importance:
        task["importance"] = importance
    if status:
        task["status"] = status
    if due_date:
        task["dueDateTime"] = _build_datetime(due_date, due_timezone)
    if start_date:
        task["startDateTime"] = _build_datetime(start_date, start_timezone)
    if reminder_date:
        task["reminderDateTime"] = _build_datetime(reminder_date, reminder_timezone)
        if is_reminder_on is None:
            is_reminder_on = True
    if is_reminder_on is not None:
        task["isReminderOn"] = is_reminder_on
    if categories:
        task["categories"] = categories

    if recurrence_pattern:
        task["recurrence"] = {
            "pattern": {
                "type": recurrence_pattern,
                "interval": recurrence_interval or 1,
            },
            "range": _build_recurrence_range(
                recurrence_range_type,
                recurrence_range_end_date,
                recurrence_occurrences,
                start_date,
            ),
        }

    result = await _graph("POST", f"/me/todo/lists/{list_id}/tasks", body=task)
    return json.dumps(result, indent=2)


@mcp.tool()
async def get_task(list_id: str, task_id: str) -> str:
    """Get a specific task by ID.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
    """
    result = await _graph("GET", f"/me/todo/lists/{list_id}/tasks/{task_id}")
    return json.dumps(result, indent=2)


@mcp.tool()
async def update_task(
    list_id: str,
    task_id: str,
    title: Optional[str] = None,
    body_content: Optional[str] = None,
    body_content_type: Optional[str] = None,
    importance: Optional[str] = None,
    status: Optional[str] = None,
    due_date: Optional[str] = None,
    due_timezone: Optional[str] = "UTC",
    reminder_date: Optional[str] = None,
    reminder_timezone: Optional[str] = "UTC",
    is_reminder_on: Optional[bool] = None,
    categories: Optional[list[str]] = None,
    start_date: Optional[str] = None,
    start_timezone: Optional[str] = "UTC",
    recurrence_pattern: Optional[str] = None,
    recurrence_interval: Optional[int] = None,
    recurrence_range_type: Optional[str] = None,
    recurrence_range_end_date: Optional[str] = None,
    recurrence_occurrences: Optional[int] = None,
) -> str:
    """Update any property of an existing task. Only supply the fields you want to change.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task to update.
        title: New title for the task.
        body_content: New body/notes content.
        body_content_type: Body format – "text" or "html".
        importance: "low", "normal", or "high".
        status: "notStarted", "inProgress", "completed", "waitingOnOthers", "deferred".
        due_date: New due date (ISO 8601). Set to empty string to clear.
        due_timezone: Time zone for the due date.
        reminder_date: New reminder date/time (ISO 8601).
        reminder_timezone: Time zone for the reminder.
        is_reminder_on: Enable or disable the reminder.
        categories: Updated list of category labels.
        start_date: New start date (ISO 8601).
        start_timezone: Time zone for the start date.
        recurrence_pattern: Recurrence type – "daily", "weekly", "absoluteMonthly",
                            "absoluteYearly", "relativeMonthly", or "relativeYearly".
        recurrence_interval: How often the pattern repeats (e.g. 1 = every week).
        recurrence_range_type: "endDate", "noEnd", or "numbered".
        recurrence_range_end_date: End date for recurrence (ISO 8601 date).
        recurrence_occurrences: Number of occurrences (used when recurrence_range_type="numbered").
    """
    patch: dict[str, Any] = {}

    if title is not None:
        patch["title"] = title
    if body_content is not None:
        patch["body"] = {"content": body_content, "contentType": body_content_type or "text"}
    if importance is not None:
        patch["importance"] = importance
    if status is not None:
        patch["status"] = status
    if due_date is not None:
        patch["dueDateTime"] = _build_datetime(due_date, due_timezone) if due_date else None
    if start_date is not None:
        patch["startDateTime"] = _build_datetime(start_date, start_timezone) if start_date else None
    if reminder_date is not None:
        patch["reminderDateTime"] = _build_datetime(reminder_date, reminder_timezone) if reminder_date else None
    if is_reminder_on is not None:
        patch["isReminderOn"] = is_reminder_on
    if categories is not None:
        patch["categories"] = categories
    if recurrence_pattern is not None:
        patch["recurrence"] = {
            "pattern": {
                "type": recurrence_pattern,
                "interval": recurrence_interval or 1,
            },
            "range": _build_recurrence_range(
                recurrence_range_type,
                recurrence_range_end_date,
                recurrence_occurrences,
                start_date,
            ),
        }

    if not patch:
        return json.dumps({"error": "No properties provided to update."})

    result = await _graph("PATCH", f"/me/todo/lists/{list_id}/tasks/{task_id}", body=patch)
    return json.dumps(result, indent=2)


@mcp.tool()
async def delete_task(list_id: str, task_id: str) -> str:
    """Delete a task permanently.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task to delete.
    """
    result = await _graph("DELETE", f"/me/todo/lists/{list_id}/tasks/{task_id}")
    return json.dumps(result, indent=2)


@mcp.tool()
async def complete_task(list_id: str, task_id: str) -> str:
    """Mark a task as completed. Convenience shortcut for update_task with status=completed.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task to complete.
    """
    result = await _graph(
        "PATCH",
        f"/me/todo/lists/{list_id}/tasks/{task_id}",
        body={"status": "completed"},
    )
    return json.dumps(result, indent=2)


# ═══════════════════════════════════════════════════════════
#  3. CHECKLIST ITEMS (STEPS)
# ═══════════════════════════════════════════════════════════

@mcp.tool()
async def list_checklist_items(list_id: str, task_id: str) -> str:
    """List all checklist items (steps) on a task.

    In the To Do app these appear as "Steps" under a task.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
    """
    result = await _graph_paged(f"/me/todo/lists/{list_id}/tasks/{task_id}/checklistItems")
    return json.dumps(result, indent=2)


@mcp.tool()
async def create_checklist_item(list_id: str, task_id: str, display_name: str) -> str:
    """Add a new checklist item (step) to a task.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        display_name: The text of the step, e.g. "Draft the letter".
    """
    result = await _graph(
        "POST",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/checklistItems",
        body={"displayName": display_name},
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def get_checklist_item(list_id: str, task_id: str, checklist_item_id: str) -> str:
    """Get a specific checklist item (step) by ID.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        checklist_item_id: The ID of the checklist item.
    """
    result = await _graph(
        "GET",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/checklistItems/{checklist_item_id}",
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def update_checklist_item(
    list_id: str,
    task_id: str,
    checklist_item_id: str,
    display_name: Optional[str] = None,
    is_checked: Optional[bool] = None,
) -> str:
    """Update a checklist item (step) – rename it or toggle its checked state.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        checklist_item_id: The ID of the checklist item.
        display_name: New text for the step (optional).
        is_checked: Set True to check / False to uncheck the step (optional).
    """
    patch: dict[str, Any] = {}
    if display_name is not None:
        patch["displayName"] = display_name
    if is_checked is not None:
        patch["isChecked"] = is_checked

    if not patch:
        return json.dumps({"error": "No properties provided to update."})

    result = await _graph(
        "PATCH",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/checklistItems/{checklist_item_id}",
        body=patch,
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def delete_checklist_item(list_id: str, task_id: str, checklist_item_id: str) -> str:
    """Delete a checklist item (step) from a task.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        checklist_item_id: The ID of the checklist item to delete.
    """
    result = await _graph(
        "DELETE",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/checklistItems/{checklist_item_id}",
    )
    return json.dumps(result, indent=2)


# ═══════════════════════════════════════════════════════════
#  4. LINKED RESOURCES
# ═══════════════════════════════════════════════════════════

@mcp.tool()
async def list_linked_resources(list_id: str, task_id: str) -> str:
    """List all linked resources on a task.

    Linked resources connect a task back to an external item
    (e.g. the email or document that spawned the task).

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
    """
    result = await _graph_paged(f"/me/todo/lists/{list_id}/tasks/{task_id}/linkedResources")
    return json.dumps(result, indent=2)


@mcp.tool()
async def create_linked_resource(
    list_id: str,
    task_id: str,
    web_url: str,
    application_name: Optional[str] = None,
    display_name: Optional[str] = None,
    external_id: Optional[str] = None,
) -> str:
    """Link an external item to a task.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        web_url: Deep link URL to the external item.
        application_name: Name of the source app (e.g. "Outlook", "Matter 1234").
        display_name: Display title for the link.
        external_id: ID of the item in the external system.
    """
    body: dict[str, Any] = {"webUrl": web_url}
    if application_name:
        body["applicationName"] = application_name
    if display_name:
        body["displayName"] = display_name
    if external_id:
        body["externalId"] = external_id

    result = await _graph(
        "POST",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/linkedResources",
        body=body,
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def get_linked_resource(list_id: str, task_id: str, linked_resource_id: str) -> str:
    """Get a specific linked resource by ID.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        linked_resource_id: The ID of the linked resource.
    """
    result = await _graph(
        "GET",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/linkedResources/{linked_resource_id}",
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def update_linked_resource(
    list_id: str,
    task_id: str,
    linked_resource_id: str,
    web_url: Optional[str] = None,
    application_name: Optional[str] = None,
    display_name: Optional[str] = None,
    external_id: Optional[str] = None,
) -> str:
    """Update a linked resource.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        linked_resource_id: The ID of the linked resource.
        web_url: New deep link URL.
        application_name: New application name.
        display_name: New display name.
        external_id: New external ID.
    """
    patch: dict[str, Any] = {}
    if web_url is not None:
        patch["webUrl"] = web_url
    if application_name is not None:
        patch["applicationName"] = application_name
    if display_name is not None:
        patch["displayName"] = display_name
    if external_id is not None:
        patch["externalId"] = external_id

    if not patch:
        return json.dumps({"error": "No properties provided to update."})

    result = await _graph(
        "PATCH",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/linkedResources/{linked_resource_id}",
        body=patch,
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def delete_linked_resource(list_id: str, task_id: str, linked_resource_id: str) -> str:
    """Delete a linked resource from a task.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        linked_resource_id: The ID of the linked resource to delete.
    """
    result = await _graph(
        "DELETE",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/linkedResources/{linked_resource_id}",
    )
    return json.dumps(result, indent=2)


# ═══════════════════════════════════════════════════════════
#  5. ATTACHMENTS (beta)
# ═══════════════════════════════════════════════════════════

@mcp.tool()
async def list_attachments(list_id: str, task_id: str) -> str:
    """List all file attachments on a task.

    NOTE: Uses the Microsoft Graph beta endpoint.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
    """
    result = await _graph_paged(
        f"/me/todo/lists/{list_id}/tasks/{task_id}/attachments",
        base=GRAPH_BETA,
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def create_attachment(
    list_id: str,
    task_id: str,
    name: str,
    content_type: str,
    content_bytes_base64: str,
) -> str:
    """Upload a file attachment to a task.

    NOTE: Uses the Microsoft Graph beta endpoint.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        name: File name including extension, e.g. "brief.pdf".
        content_type: MIME type, e.g. "application/pdf".
        content_bytes_base64: The file content encoded as a base64 string.
    """
    body = {
        "@odata.type": "#microsoft.graph.taskFileAttachment",
        "name": name,
        "contentType": content_type,
        "contentBytes": content_bytes_base64,
    }
    result = await _graph(
        "POST",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/attachments",
        body=body,
        base=GRAPH_BETA,
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def get_attachment(list_id: str, task_id: str, attachment_id: str) -> str:
    """Get a specific file attachment by ID.

    NOTE: Uses the Microsoft Graph beta endpoint.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        attachment_id: The ID of the attachment.
    """
    result = await _graph(
        "GET",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/attachments/{attachment_id}",
        base=GRAPH_BETA,
    )
    return json.dumps(result, indent=2)


@mcp.tool()
async def delete_attachment(list_id: str, task_id: str, attachment_id: str) -> str:
    """Delete a file attachment from a task.

    NOTE: Uses the Microsoft Graph beta endpoint.

    Args:
        list_id: The ID of the task list.
        task_id: The ID of the task.
        attachment_id: The ID of the attachment to delete.
    """
    result = await _graph(
        "DELETE",
        f"/me/todo/lists/{list_id}/tasks/{task_id}/attachments/{attachment_id}",
        base=GRAPH_BETA,
    )
    return json.dumps(result, indent=2)


# ═══════════════════════════════════════════════════════════
#  6. DELTA / SYNC
# ═══════════════════════════════════════════════════════════

@mcp.tool()
async def delta_tasks(list_id: str, delta_token: Optional[str] = None) -> str:
    """Get incremental changes to tasks since the last sync (delta query).

    On the first call, omit delta_token to get all tasks plus a deltaLink.
    On subsequent calls, pass the deltaToken from the previous response to
    get only what changed since then. This is much more efficient than
    re-fetching the full list for sync-style workflows.

    Args:
        list_id: The ID of the task list.
        delta_token: Token from a previous delta response (@odata.deltaLink).
                     Omit for the initial full sync.
    """
    if delta_token:
        path = f"/me/todo/lists/{list_id}/tasks/delta?$deltatoken={delta_token}"
    else:
        path = f"/me/todo/lists/{list_id}/tasks/delta"

    token = await _get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url: str | None = f"{GRAPH_V1}{path}"
    all_values: list[Any] = []
    delta_link: str | None = None

    async with httpx.AsyncClient(timeout=30) as client:
        while url:
            log.info("GET %s (delta)", url)
            resp = await client.get(url, headers=headers)
            if resp.status_code >= 400:
                try:
                    err = resp.json()
                except Exception:
                    err = resp.text
                return json.dumps({"error": True, "status_code": resp.status_code, "detail": err})
            data = resp.json()
            all_values.extend(data.get("value", []))
            delta_link = data.get("@odata.deltaLink")
            url = data.get("@odata.nextLink")

    next_delta_token = None
    if delta_link and "$deltatoken=" in delta_link:
        next_delta_token = delta_link.split("$deltatoken=")[-1]

    return json.dumps(
        {
            "value": all_values,
            "@odata.count": len(all_values),
            "@odata.deltaLink": delta_link,
            "deltaToken": next_delta_token,
        },
        indent=2,
    )


# ═══════════════════════════════════════════════════════════
#  7. CONVENIENCE / SMART TOOLS
# ═══════════════════════════════════════════════════════════

@mcp.tool()
async def find_list_by_name(name: str) -> str:
    """Find a task list by name (case-insensitive partial match).

    Use this to resolve a human-readable list name like "Next Actions"
    to the list ID needed by other tools.

    Args:
        name: The name (or partial name) to search for.
    """
    result = await _graph_paged("/me/todo/lists")
    if "value" in result:
        matches = [
            lst for lst in result["value"]
            if name.lower() in lst.get("displayName", "").lower()
        ]
        if matches:
            return json.dumps({"found": len(matches), "lists": matches}, indent=2)
        return json.dumps(
            {
                "found": 0,
                "message": f"No list matching '{name}'. Available lists: "
                + ", ".join(lst.get("displayName", "?") for lst in result["value"]),
            }
        )
    return json.dumps(result, indent=2)


@mcp.tool()
async def find_task_by_title(list_id: str, title: str) -> str:
    """Find a task by title (case-insensitive partial match) in a list.
    Searches all pages of results.

    Use this to resolve a human-readable task title to the task ID
    needed by other tools.

    Args:
        list_id: The ID of the task list to search in.
        title: The title (or partial title) to search for.
    """
    result = await _graph_paged(f"/me/todo/lists/{list_id}/tasks")
    if "value" in result:
        matches = [
            t for t in result["value"]
            if title.lower() in t.get("title", "").lower()
        ]
        if matches:
            return json.dumps({"found": len(matches), "tasks": matches}, indent=2)
        return json.dumps({"found": 0, "message": f"No task matching '{title}' in this list."})
    return json.dumps(result, indent=2)


# ═══════════════════════════════════════════════════════════
#  RECURRENCE HELPER (internal)
# ═══════════════════════════════════════════════════════════

def _build_recurrence_range(
    range_type: Optional[str],
    end_date: Optional[str],
    occurrences: Optional[int],
    start_date: Optional[str],
) -> dict[str, Any]:
    rng: dict[str, Any] = {
        "type": range_type or "noEnd",
        "startDate": (start_date or "").split("T")[0] if start_date else "",
    }
    if range_type == "endDate" and end_date:
        rng["endDate"] = end_date
    if range_type == "numbered" and occurrences:
        rng["numberOfOccurrences"] = occurrences
    return rng


# ═══════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════

if __name__ == "__main__":
    # MCP_TRANSPORT=stdio  → for Claude Desktop (local, stdio)
    # Otherwise            → streamable HTTP (for Copilot Studio, remote clients)
    transport = os.environ.get("MCP_TRANSPORT", "streamable-http")

    if transport == "stdio":
        mcp.run(transport="stdio")
    else:
        port = int(os.environ.get("PORT", 3000))
        log.info("Starting MCP server on http://0.0.0.0:%d/mcp", port)
        mcp.run(transport="streamable-http", host="0.0.0.0", port=port)
