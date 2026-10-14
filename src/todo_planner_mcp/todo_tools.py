"""Microsoft To Do tools (lists = projects, tasks, checklist items)."""

from __future__ import annotations

from typing import Any, Literal
from urllib.parse import quote
from zoneinfo import ZoneInfo

from .app import DESTRUCTIVE, READ_ONLY, WRITE, ctx, pick, resolve_by_name, tool, user_tz
from .dates import todo_datetime

Status = Literal["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"]
Importance = Literal["low", "normal", "high"]

_TASK_FIELDS = (
    "id", "title", "status", "importance", "dueDateTime", "startDateTime", "reminderDateTime",
    "isReminderOn", "completedDateTime", "categories", "createdDateTime", "lastModifiedDateTime",
)


def _q(segment: str) -> str:
    return quote(segment, safe="")


def _shape_task(t: dict[str, Any], include_body: bool = False) -> dict[str, Any]:
    out = pick(t, *_TASK_FIELDS)
    for key in ("dueDateTime", "startDateTime", "reminderDateTime", "completedDateTime"):
        if key in out:
            out[key] = f"{out[key].get('dateTime', '')[:16]} {out[key].get('timeZone', '')}".strip()
    if include_body and (t.get("body") or {}).get("content"):
        out["notes"] = t["body"]["content"]
    if t.get("checklistItems"):
        out["checklist"] = [pick(c, "id", "displayName", "isChecked") for c in t["checklistItems"]]
    return out


async def _lists() -> list[dict]:
    return await ctx().graph.get_all("/me/todo/lists")


async def _list_id(list_ref: str) -> str:
    return resolve_by_name(await _lists(), list_ref, "displayName", "To Do list")["id"]


def _task_body(
    tz: ZoneInfo, title: str | None, notes: str | None, due: str | None, start: str | None, reminder: str | None,
    importance: str | None, status: str | None, categories: list[str] | None,
) -> dict[str, Any]:
    body: dict[str, Any] = {}
    if title is not None:
        if not title.strip():
            raise ValueError("title cannot be empty")
        body["title"] = title.strip()
    if notes is not None:
        body["body"] = {"content": notes, "contentType": "text"}
    # Empty string clears a date.
    if due is not None:
        body["dueDateTime"] = todo_datetime(due, tz) if due else None
    if start is not None:
        body["startDateTime"] = todo_datetime(start, tz) if start else None
    if reminder is not None:
        body["reminderDateTime"] = todo_datetime(reminder, tz) if reminder else None
        body["isReminderOn"] = bool(reminder)
    if importance is not None:
        body["importance"] = importance
    if status is not None:
        body["status"] = status
    if categories is not None:
        body["categories"] = categories
    return body


# ---------------------------------------------------------------- lists

@tool(READ_ONLY)
async def todo_list_lists() -> list[dict]:
    """List all Microsoft To Do task lists (projects) for the signed-in user."""
    return [pick(l, "id", "displayName", "wellknownListName", "isOwner", "isShared") for l in await _lists()]


@tool(WRITE)
async def todo_create_list(name: str) -> dict:
    """Create a new Microsoft To Do task list (project)."""
    if not name.strip():
        raise ValueError("name cannot be empty")
    created = await ctx().graph.post("/me/todo/lists", {"displayName": name.strip()})
    return pick(created, "id", "displayName")


@tool(WRITE)
async def todo_rename_list(list_ref: str, new_name: str) -> dict:
    """Rename a To Do list. list_ref is the list ID or its exact name."""
    list_id = await _list_id(list_ref)
    updated = await ctx().graph.patch(f"/me/todo/lists/{_q(list_id)}", {"displayName": new_name.strip()})
    return pick(updated, "id", "displayName") or {"id": list_id, "displayName": new_name}


@tool(DESTRUCTIVE)
async def todo_delete_list(list_ref: str) -> dict:
    """Permanently delete a To Do list AND all tasks in it. Confirm with the user first.
    The default 'Tasks' list cannot be deleted."""
    target = resolve_by_name(await _lists(), list_ref, "displayName", "To Do list")
    if target.get("wellknownListName") not in (None, "none"):
        raise ValueError(f"'{target['displayName']}' is a built-in list and cannot be deleted.")
    await ctx().graph.delete(f"/me/todo/lists/{_q(target['id'])}")
    return {"deleted": True, "id": target["id"], "displayName": target["displayName"]}


# ---------------------------------------------------------------- tasks

@tool(READ_ONLY)
async def todo_list_tasks(list_ref: str, include_completed: bool = False, max_items: int = 100) -> list[dict]:
    """List tasks in a To Do list. list_ref is the list ID or exact name (e.g. 'Tasks')."""
    list_id = await _list_id(list_ref)
    params: dict[str, Any] = {"$top": min(max_items, 100)}
    if not include_completed:
        params["$filter"] = "status ne 'completed'"
    tasks = await ctx().graph.get_all(f"/me/todo/lists/{_q(list_id)}/tasks", params, max_items=max_items)
    return [_shape_task(t) for t in tasks]


@tool(READ_ONLY)
async def todo_get_task(list_ref: str, task_id: str) -> dict:
    """Get one To Do task with notes and checklist items."""
    list_id = await _list_id(list_ref)
    task = await ctx().graph.get(
        f"/me/todo/lists/{_q(list_id)}/tasks/{_q(task_id)}", {"$expand": "checklistItems"}
    )
    return _shape_task(task, include_body=True)


@tool(WRITE)
async def todo_create_task(
    list_ref: str,
    title: str,
    notes: str | None = None,
    due: str | None = None,
    start: str | None = None,
    reminder: str | None = None,
    importance: Importance | None = None,
    categories: list[str] | None = None,
    checklist: list[str] | None = None,
) -> dict:
    """Create a task in a To Do list. Dates: YYYY-MM-DD or YYYY-MM-DDTHH:MM (local time).
    checklist adds sub-steps."""
    list_id = await _list_id(list_ref)
    body = _task_body(await user_tz(), title, notes, due, start, reminder, importance, None, categories)
    if checklist:
        body["checklistItems"] = [{"displayName": c} for c in checklist if c.strip()]
    created = await ctx().graph.post(f"/me/todo/lists/{_q(list_id)}/tasks", body)
    return _shape_task(created, include_body=True)


@tool(WRITE)
async def todo_update_task(
    list_ref: str,
    task_id: str,
    title: str | None = None,
    notes: str | None = None,
    due: str | None = None,
    start: str | None = None,
    reminder: str | None = None,
    importance: Importance | None = None,
    status: Status | None = None,
    categories: list[str] | None = None,
) -> dict:
    """Update a To Do task. Only supplied fields change. Pass an empty string for a date to clear it.
    Set status='completed' to complete a task, 'notStarted' to reopen it."""
    list_id = await _list_id(list_ref)
    body = _task_body(await user_tz(), title, notes, due, start, reminder, importance, status, categories)
    if not body:
        raise ValueError("Nothing to update: supply at least one field.")
    updated = await ctx().graph.patch(f"/me/todo/lists/{_q(list_id)}/tasks/{_q(task_id)}", body)
    return _shape_task(updated or {"id": task_id, **body}, include_body=True)


@tool(DESTRUCTIVE)
async def todo_delete_task(list_ref: str, task_id: str) -> dict:
    """Permanently delete a To Do task. Confirm with the user first."""
    list_id = await _list_id(list_ref)
    await ctx().graph.delete(f"/me/todo/lists/{_q(list_id)}/tasks/{_q(task_id)}")
    return {"deleted": True, "id": task_id}


@tool(WRITE)
async def todo_move_task(from_list_ref: str, task_id: str, to_list_ref: str) -> dict:
    """Move a task to another To Do list (copies title, notes, dates, importance, status,
    categories and checklist, then deletes the original). The task gets a new ID."""
    src = await _list_id(from_list_ref)
    dst = await _list_id(to_list_ref)
    g = ctx().graph
    task = await g.get(f"/me/todo/lists/{_q(src)}/tasks/{_q(task_id)}", {"$expand": "checklistItems"})
    body = {k: task[k] for k in ("title", "body", "dueDateTime", "startDateTime", "reminderDateTime",
                                 "isReminderOn", "importance", "status", "categories") if task.get(k) is not None}
    if task.get("checklistItems"):
        body["checklistItems"] = [pick(c, "displayName", "isChecked") for c in task["checklistItems"]]
    created = await g.post(f"/me/todo/lists/{_q(dst)}/tasks", body)
    await g.delete(f"/me/todo/lists/{_q(src)}/tasks/{_q(task_id)}")
    return _shape_task(created)


# ---------------------------------------------------------------- checklist

@tool(WRITE)
async def todo_add_checklist_item(list_ref: str, task_id: str, text: str) -> dict:
    """Add a checklist (sub-step) item to a To Do task."""
    list_id = await _list_id(list_ref)
    created = await ctx().graph.post(
        f"/me/todo/lists/{_q(list_id)}/tasks/{_q(task_id)}/checklistItems", {"displayName": text}
    )
    return pick(created, "id", "displayName", "isChecked")


@tool(WRITE)
async def todo_update_checklist_item(
    list_ref: str, task_id: str, item_id: str, text: str | None = None, checked: bool | None = None
) -> dict:
    """Rename and/or tick/untick a checklist item on a To Do task."""
    body: dict[str, Any] = {}
    if text is not None:
        body["displayName"] = text
    if checked is not None:
        body["isChecked"] = checked
    if not body:
        raise ValueError("Nothing to update.")
    list_id = await _list_id(list_ref)
    updated = await ctx().graph.patch(
        f"/me/todo/lists/{_q(list_id)}/tasks/{_q(task_id)}/checklistItems/{_q(item_id)}", body
    )
    return pick(updated, "id", "displayName", "isChecked")


@tool(DESTRUCTIVE)
async def todo_delete_checklist_item(list_ref: str, task_id: str, item_id: str) -> dict:
    """Delete a checklist item from a To Do task. Confirm with the user first."""
    list_id = await _list_id(list_ref)
    await ctx().graph.delete(f"/me/todo/lists/{_q(list_id)}/tasks/{_q(task_id)}/checklistItems/{_q(item_id)}")
    return {"deleted": True, "id": item_id}
