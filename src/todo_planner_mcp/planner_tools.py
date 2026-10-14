"""Microsoft Planner tools (plans = projects, buckets, tasks, task details).

Planner uses optimistic concurrency: every update/delete needs the item's current
ETag in If-Match. These tools fetch the latest ETag automatically just before writing.
"""

from __future__ import annotations

import asyncio
import uuid
from typing import Any, Literal
from urllib.parse import quote

from .app import DESTRUCTIVE, READ_ONLY, WRITE, ctx, looks_like_id, pick, resolve_by_name, tool, user_tz
from .dates import planner_datetime
from .graph import GraphError

Priority = Literal["urgent", "important", "medium", "low"]
Progress = Literal["notStarted", "inProgress", "completed"]

_PRIORITY_TO_INT = {"urgent": 1, "important": 3, "medium": 5, "low": 9}
_PROGRESS_TO_PCT = {"notStarted": 0, "inProgress": 50, "completed": 100}


def _q(segment: str) -> str:
    return quote(segment, safe="")


def _etag(obj: dict) -> str:
    tag = obj.get("@odata.etag")
    if not tag:
        raise RuntimeError("Graph did not return an ETag for this Planner item.")
    return tag


def _priority_name(value: int | None) -> str | None:
    if value is None:
        return None
    if value <= 1:
        return "urgent"
    if value <= 4:
        return "important"
    if value <= 7:
        return "medium"
    return "low"


def _progress_name(pct: int | None) -> str:
    return "completed" if pct == 100 else "inProgress" if pct else "notStarted"


async def _me_id() -> str:
    return (await ctx().graph.get("/me", {"$select": "id"}))["id"]


async def _resolve_user_ids(refs: list[str]) -> list[str]:
    """'me', a user object ID, or a UPN/email. Email lookup needs User.ReadBasic.All."""
    ids: list[str] = []
    for ref in refs:
        ref = ref.strip()
        if ref.lower() == "me":
            ids.append(await _me_id())
        elif "@" in ref:
            try:
                ids.append((await ctx().graph.get(f"/users/{_q(ref)}", {"$select": "id"}))["id"])
            except GraphError as exc:
                if exc.status == 403:
                    raise PermissionError(
                        "Looking up users by email needs the User.ReadBasic.All permission. Add it to "
                        "TODO_MCP_EXTRA_SCOPES and sign in again, or pass the user's object ID."
                    ) from exc
                raise
        else:
            ids.append(ref)
    return ids


async def _plans() -> list[dict]:
    return await ctx().graph.get_all("/me/planner/plans")


async def _plan(plan_ref: str) -> dict:
    plans = await _plans()
    try:
        return resolve_by_name(plans, plan_ref, "title", "Planner plan")
    except ValueError:
        if looks_like_id(plan_ref):  # a plan the user can access but isn't listed under /me
            return await ctx().graph.get(f"/planner/plans/{_q(plan_ref)}")
        raise


async def _buckets(plan_id: str) -> list[dict]:
    return await ctx().graph.get_all(f"/planner/plans/{_q(plan_id)}/buckets")


async def _bucket_id(plan_id: str, bucket_ref: str) -> str:
    return resolve_by_name(await _buckets(plan_id), bucket_ref, "name", "bucket")["id"]


async def _label_keys(plan_id: str, labels: list[str]) -> dict[str, bool]:
    """Map label names (or 'category1'..'category25') to Planner appliedCategories."""
    details = await ctx().graph.get(f"/planner/plans/{_q(plan_id)}/details")
    names = {(v or "").lower(): k for k, v in (details.get("categoryDescriptions") or {}).items() if v}
    applied: dict[str, bool] = {}
    for label in labels:
        key = label if label.lower().startswith("category") else names.get(label.lower())
        if not key:
            raise ValueError(f"Label '{label}' is not defined in this plan. Defined: {sorted(names) or 'none'}")
        applied[key] = True
    return applied


def _shape_task(t: dict, bucket_names: dict[str, str] | None = None, me: str | None = None) -> dict:
    out = pick(t, "id", "title", "bucketId", "dueDateTime", "startDateTime", "createdDateTime", "completedDateTime")
    out["progress"] = _progress_name(t.get("percentComplete"))
    if (p := _priority_name(t.get("priority"))) is not None:
        out["priority"] = p
    if bucket_names and t.get("bucketId") in bucket_names:
        out["bucket"] = bucket_names[t["bucketId"]]
    assignees = list((t.get("assignments") or {}).keys())
    if assignees:
        out["assignees"] = ["me" if a == me else a for a in assignees]
    labels = [k for k, v in (t.get("appliedCategories") or {}).items() if v]
    if labels:
        out["labels"] = labels
    return out


async def _get_details_with_retry(task_id: str) -> dict:
    # Details can lag briefly behind task creation (eventual consistency).
    for attempt in range(4):
        try:
            return await ctx().graph.get(f"/planner/tasks/{_q(task_id)}/details")
        except GraphError as exc:
            if exc.status not in (404, 409) or attempt == 3:
                raise
            await asyncio.sleep(1 + attempt)
    raise RuntimeError("unreachable")


async def _update_details(
    task_id: str, description: str | None, add_checklist: list[str] | None, check_items: list[str] | None
) -> None:
    if description is None and not add_checklist and not check_items:
        return
    details = await _get_details_with_retry(task_id)
    body: dict[str, Any] = {}
    if description is not None:
        body["description"] = description
        body["previewType"] = "description" if description else "automatic"
    checklist: dict[str, Any] = {}
    for title in add_checklist or []:
        if title.strip():
            checklist[str(uuid.uuid4())] = {
                "@odata.type": "microsoft.graph.plannerChecklistItem", "title": title.strip(), "isChecked": False,
            }
    existing = details.get("checklist") or {}
    for ref in check_items or []:
        key = ref if ref in existing else next(
            (k for k, v in existing.items() if (v.get("title") or "").lower() == ref.lower()), None
        )
        if not key:
            raise ValueError(f"Checklist item '{ref}' not found on this task.")
        checklist[key] = {"@odata.type": "microsoft.graph.plannerChecklistItem", "isChecked": True}
    if checklist:
        body["checklist"] = checklist
    await ctx().graph.patch(f"/planner/tasks/{_q(task_id)}/details", body, etag=_etag(details))


# ---------------------------------------------------------------- groups & plans

@tool(READ_ONLY)
async def planner_list_my_groups() -> list[dict]:
    """List Microsoft 365 groups the user belongs to (needed as the container when creating a plan).
    Requires the GroupMember.Read.All permission; otherwise use group IDs from planner_list_plans."""
    try:
        groups = await ctx().graph.get_all(
            "/me/memberOf/microsoft.graph.group", {"$select": "id,displayName,groupTypes,mail"}
        )
    except GraphError as exc:
        if exc.status == 403:
            raise PermissionError(
                "Listing groups needs GroupMember.Read.All. Add it to TODO_MCP_EXTRA_SCOPES and sign in "
                "again (admin consent required), or take the groupId from planner_list_plans."
            ) from exc
        raise
    return [pick(g, "id", "displayName", "mail") for g in groups if "Unified" in (g.get("groupTypes") or [])]


@tool(READ_ONLY)
async def planner_list_plans() -> list[dict]:
    """List Planner plans (projects) shared with the signed-in user, with their group IDs."""
    return [
        {**pick(p, "id", "title", "createdDateTime"), "groupId": (p.get("container") or {}).get("containerId") or p.get("owner")}
        for p in await _plans()
    ]


@tool(WRITE)
async def planner_create_plan(group_id: str, title: str, buckets: list[str] | None = None) -> dict:
    """Create a Planner plan (project) in a Microsoft 365 group, optionally with initial buckets
    (e.g. ['To do', 'In progress', 'Done'])."""
    g = ctx().graph
    plan = await g.post("/planner/plans", {
        "container": {"url": f"https://graph.microsoft.com/v1.0/groups/{_q(group_id)}"},
        "title": title.strip(),
    })
    created_buckets = []
    order_hint = " !"
    for name in buckets or []:
        b = await g.post("/planner/buckets", {"name": name, "planId": plan["id"], "orderHint": order_hint})
        order_hint = f"{b.get('orderHint', '')} !"  # keep buckets in the order given
        created_buckets.append(pick(b, "id", "name"))
    return {**pick(plan, "id", "title"), "groupId": group_id, "buckets": created_buckets}


@tool(WRITE)
async def planner_rename_plan(plan_ref: str, new_title: str) -> dict:
    """Rename a Planner plan. plan_ref is the plan ID or exact title."""
    plan = await _plan(plan_ref)
    fresh = await ctx().graph.get(f"/planner/plans/{_q(plan['id'])}")
    updated = await ctx().graph.patch(f"/planner/plans/{_q(plan['id'])}", {"title": new_title.strip()}, etag=_etag(fresh))
    return pick(updated, "id", "title")


@tool(DESTRUCTIVE)
async def planner_delete_plan(plan_ref: str) -> dict:
    """Permanently delete a Planner plan with ALL its buckets and tasks. Confirm with the user first."""
    plan = await _plan(plan_ref)
    fresh = await ctx().graph.get(f"/planner/plans/{_q(plan['id'])}")
    await ctx().graph.delete(f"/planner/plans/{_q(plan['id'])}", etag=_etag(fresh))
    return {"deleted": True, "id": plan["id"], "title": plan.get("title")}


# ---------------------------------------------------------------- buckets

@tool(READ_ONLY)
async def planner_list_buckets(plan_ref: str) -> list[dict]:
    """List buckets (columns) in a Planner plan."""
    plan = await _plan(plan_ref)
    buckets = sorted(await _buckets(plan["id"]), key=lambda b: b.get("orderHint", ""))
    return [pick(b, "id", "name") for b in buckets]


@tool(WRITE)
async def planner_create_bucket(plan_ref: str, name: str) -> dict:
    """Create a bucket in a Planner plan."""
    plan = await _plan(plan_ref)
    b = await ctx().graph.post("/planner/buckets", {"name": name.strip(), "planId": plan["id"], "orderHint": " !"})
    return pick(b, "id", "name")


@tool(WRITE)
async def planner_rename_bucket(plan_ref: str, bucket_ref: str, new_name: str) -> dict:
    """Rename a bucket in a Planner plan."""
    plan = await _plan(plan_ref)
    bucket_id = await _bucket_id(plan["id"], bucket_ref)
    fresh = await ctx().graph.get(f"/planner/buckets/{_q(bucket_id)}")
    updated = await ctx().graph.patch(f"/planner/buckets/{_q(bucket_id)}", {"name": new_name.strip()}, etag=_etag(fresh))
    return pick(updated, "id", "name")


@tool(DESTRUCTIVE)
async def planner_delete_bucket(plan_ref: str, bucket_ref: str) -> dict:
    """Permanently delete a bucket and ALL tasks in it. Confirm with the user first."""
    plan = await _plan(plan_ref)
    bucket_id = await _bucket_id(plan["id"], bucket_ref)
    fresh = await ctx().graph.get(f"/planner/buckets/{_q(bucket_id)}")
    await ctx().graph.delete(f"/planner/buckets/{_q(bucket_id)}", etag=_etag(fresh))
    return {"deleted": True, "id": bucket_id}


# ---------------------------------------------------------------- tasks

@tool(READ_ONLY)
async def planner_list_tasks(
    plan_ref: str, bucket_ref: str | None = None, include_completed: bool = False, assigned_to_me: bool = False
) -> list[dict]:
    """List tasks in a Planner plan, optionally filtered by bucket, completion and assignment."""
    plan = await _plan(plan_ref)
    buckets = await _buckets(plan["id"])
    names = {b["id"]: b["name"] for b in buckets}
    tasks = await ctx().graph.get_all(f"/planner/plans/{_q(plan['id'])}/tasks")
    me = await _me_id()
    if bucket_ref:
        bid = resolve_by_name(buckets, bucket_ref, "name", "bucket")["id"]
        tasks = [t for t in tasks if t.get("bucketId") == bid]
    if not include_completed:
        tasks = [t for t in tasks if t.get("percentComplete") != 100]
    if assigned_to_me:
        tasks = [t for t in tasks if me in (t.get("assignments") or {})]
    return [_shape_task(t, names, me) for t in tasks]


@tool(READ_ONLY)
async def planner_my_tasks(include_completed: bool = False) -> list[dict]:
    """List Planner tasks assigned to the signed-in user across all plans."""
    tasks = await ctx().graph.get_all("/me/planner/tasks")
    if not include_completed:
        tasks = [t for t in tasks if t.get("percentComplete") != 100]
    titles = {p["id"]: p.get("title") for p in await _plans()}
    me = await _me_id()
    return [{**_shape_task(t, me=me), "plan": titles.get(t.get("planId"), t.get("planId"))} for t in tasks]


@tool(READ_ONLY)
async def planner_get_task(task_id: str) -> dict:
    """Get a Planner task including description and checklist."""
    g = ctx().graph
    task = await g.get(f"/planner/tasks/{_q(task_id)}")
    details = await g.get(f"/planner/tasks/{_q(task_id)}/details")
    out = _shape_task(task, me=await _me_id())
    out["planId"] = task.get("planId")
    if details.get("description"):
        out["description"] = details["description"]
    if details.get("checklist"):
        out["checklist"] = [{"id": k, "title": v.get("title"), "isChecked": v.get("isChecked")}
                            for k, v in details["checklist"].items()]
    return out


@tool(WRITE)
async def planner_create_task(
    plan_ref: str,
    title: str,
    bucket_ref: str | None = None,
    description: str | None = None,
    due: str | None = None,
    start: str | None = None,
    priority: Priority | None = None,
    progress: Progress | None = None,
    assignees: list[str] | None = None,
    labels: list[str] | None = None,
    checklist: list[str] | None = None,
) -> dict:
    """Create a Planner task. assignees: 'me', user object IDs, or emails. labels: label names
    defined on the plan. Dates: YYYY-MM-DD (local time)."""
    if not title.strip():
        raise ValueError("title cannot be empty")
    tz = await user_tz()
    plan = await _plan(plan_ref)
    body: dict[str, Any] = {"planId": plan["id"], "title": title.strip()}
    if bucket_ref:
        body["bucketId"] = await _bucket_id(plan["id"], bucket_ref)
    if due:
        body["dueDateTime"] = planner_datetime(due, tz)
    if start:
        body["startDateTime"] = planner_datetime(start, tz)
    if priority:
        body["priority"] = _PRIORITY_TO_INT[priority]
    if progress:
        body["percentComplete"] = _PROGRESS_TO_PCT[progress]
    if assignees:
        body["assignments"] = {
            uid: {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !"}
            for uid in await _resolve_user_ids(assignees)
        }
    if labels:
        body["appliedCategories"] = await _label_keys(plan["id"], labels)
    task = await ctx().graph.post("/planner/tasks", body)
    await _update_details(task["id"], description, checklist, None)
    return await planner_get_task(task["id"])


@tool(WRITE)
async def planner_update_task(
    task_id: str,
    title: str | None = None,
    bucket_ref: str | None = None,
    description: str | None = None,
    due: str | None = None,
    start: str | None = None,
    priority: Priority | None = None,
    progress: Progress | None = None,
    add_assignees: list[str] | None = None,
    remove_assignees: list[str] | None = None,
    labels: list[str] | None = None,
    add_checklist: list[str] | None = None,
    check_items: list[str] | None = None,
) -> dict:
    """Update a Planner task. Only supplied fields change. Empty string for due/start clears it.
    bucket_ref moves it to another bucket in the same plan. labels replaces the task's labels.
    check_items ticks existing checklist items (by ID or title). progress='completed' completes it."""
    g = ctx().graph
    tz = await user_tz()
    task = await g.get(f"/planner/tasks/{_q(task_id)}")
    body: dict[str, Any] = {}
    if title is not None:
        body["title"] = title.strip()
    if bucket_ref:
        body["bucketId"] = await _bucket_id(task["planId"], bucket_ref)
    if due is not None:
        body["dueDateTime"] = planner_datetime(due, tz) if due else None
    if start is not None:
        body["startDateTime"] = planner_datetime(start, tz) if start else None
    if priority:
        body["priority"] = _PRIORITY_TO_INT[priority]
    if progress:
        body["percentComplete"] = _PROGRESS_TO_PCT[progress]
    assignments: dict[str, Any] = {}
    for uid in await _resolve_user_ids(add_assignees or []):
        assignments[uid] = {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !"}
    for uid in await _resolve_user_ids(remove_assignees or []):
        assignments[uid] = None
    if assignments:
        body["assignments"] = assignments
    if labels is not None:
        wanted = await _label_keys(task["planId"], labels) if labels else {}
        current = {k for k, v in (task.get("appliedCategories") or {}).items() if v}
        body["appliedCategories"] = {**{k: False for k in current}, **wanted}

    if not body and description is None and not add_checklist and not check_items:
        raise ValueError("Nothing to update: supply at least one field.")
    if body:
        await g.patch(f"/planner/tasks/{_q(task_id)}", body, etag=_etag(task))
    await _update_details(task_id, description, add_checklist, check_items)
    return await planner_get_task(task_id)


@tool(DESTRUCTIVE)
async def planner_delete_task(task_id: str) -> dict:
    """Permanently delete a Planner task. Confirm with the user first."""
    task = await ctx().graph.get(f"/planner/tasks/{_q(task_id)}")
    await ctx().graph.delete(f"/planner/tasks/{_q(task_id)}", etag=_etag(task))
    return {"deleted": True, "id": task_id, "title": task.get("title")}
