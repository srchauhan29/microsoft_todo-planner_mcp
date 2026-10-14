# To Do & Planner MCP server

A local [MCP](https://modelcontextprotocol.io) server that lets Claude (Desktop chats, Projects and Cowork, or Claude Code) create, read, update and delete **Microsoft To Do** and **Microsoft Planner** items through Microsoft Graph, as the signed-in user.

| Area | "Project" maps to | Tools |
|---|---|---|
| Microsoft To Do | Task **list** | `todo_list_lists`, `todo_create_list`, `todo_rename_list`, `todo_delete_list`, `todo_list_tasks`, `todo_get_task`, `todo_create_task`, `todo_update_task`, `todo_move_task`, `todo_delete_task`, `todo_add_checklist_item`, `todo_update_checklist_item`, `todo_delete_checklist_item` |
| Microsoft Planner | **Plan** (in a Microsoft 365 group) | `planner_list_my_groups`, `planner_list_plans`, `planner_create_plan`, `planner_rename_plan`, `planner_delete_plan`, `planner_list_buckets`, `planner_create_bucket`, `planner_rename_bucket`, `planner_delete_bucket`, `planner_list_tasks`, `planner_my_tasks`, `planner_get_task`, `planner_create_task`, `planner_update_task`, `planner_delete_task` |
| Account | – | `auth_status`, `login`, `auto_login` |

Lists, plans and buckets can be referenced by ID **or exact name** (ambiguous names are rejected, never guessed). Delete tools are marked destructive so Claude asks before running them.

> Scope note: Graph exposes **basic Planner** plans. Planner Premium / Project for the web "projects" are Dataverse-based and are **not** reachable through these Graph endpoints.

## Quick install (any Windows machine with Claude)

Pick **one** method per machine. Installing both gives you duplicate tools.

### Option A: Claude Desktop extension (`.mcpb`), no terminal needed after uv is installed

1. Install [uv](https://docs.astral.sh/uv/) once. It's per-user, needs no admin rights and isn't an MSI:
   ```bash
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   ```
2. Download `todo-planner-<version>.mcpb` from this repo's **Releases** page, or build it yourself (see below).
3. Double-click the file, or go to Claude Desktop → **Settings → Extensions** and drag the file in. Review the optional settings (tenant, time zone, auto sign-in) and click **Install**.
4. In any chat, Project or Cowork task, say **"Log in to To Do"**. Claude asks which sign-in option you want and opens your browser.

The first start takes a little longer while uv downloads Python and the pinned dependencies. Extension settings can be changed later under Settings → Extensions.

### Option B: `install.ps1`, also works for Claude Code and Intune

From a **normal** PowerShell or Windows Terminal window (not a terminal inside Claude), in a copy of this repo:

```bash
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

The script installs uv if it's missing (it asks first), installs Python and the pinned dependencies, runs the sign-in menu (option 1 or 2), registers the server with Claude Desktop (asking you to quit Claude first) and verifies everything. Useful switches:

| Switch | Effect |
|---|---|
| `-App microsoft -Tenant contoso.com.au` | Skip the menu: Microsoft-published app, restricted to your tenant |
| `-App own -ClientId <guid> -Tenant <tenant>` | Skip the menu: your own app registration |
| `-TimeZone Australia/Sydney` | Pin a time zone (default `auto`) |
| `-ClaudeCode` | Also register with Claude Code (`claude mcp add`, user scope) |
| `-SkipLogin` / `-SkipClaudeDesktop` | Skip those steps (e.g. for an Intune install; the user then signs in from Claude) |
| `-Yes` | No confirmation prompts |
| `-Uninstall` | Remove the Claude registrations, sign out and optionally delete saved settings |

For an **Intune Win32 app**, run `install.ps1 -SkipLogin -Yes` in the **user** context. Each user then says "Log in to To Do" in Claude.

### Building the `.mcpb`

```bash
uv run python scripts\build_mcpb.py --schema path\to\mcpb-manifest-v0.4.schema.json
```

This generates `manifest.json` from the live tool list and `pyproject.toml`, validates it against the [official MCPB schema](https://github.com/modelcontextprotocol/mcpb/tree/main/schemas) when `--schema` is given (needs `uv run --with jsonschema`), and writes `dist/todo-planner-<version>.mcpb`. The bundle is built from an explicit allow-list: source, `pyproject.toml`, `uv.lock`, README and icon. It never contains `.venv`, settings or tokens. Dependencies are installed with `--frozen` from `uv.lock`.

## Security design

- **Delegated permissions only** – everything runs as you; you can only touch lists/plans you already have access to.
- **No secrets** – public client (no client secret or certificate). Only the client ID, tenant and time zone are saved, and none of them are secrets.
- **Encrypted token cache** – MSAL tokens are stored with `msal-extensions`, encrypted with Windows DPAPI under your user profile (`%USERPROFILE%\.todo-planner-mcp\msal_token_cache.bin`). Nothing is written to the repo.
- **MFA / Conditional Access honoured** – sign-in happens in the system browser against your tenant.
- **Least privilege** – default scopes are `User.Read` + `Tasks.ReadWrite`. Extra scopes are opt-in.
- Bearer tokens are only ever sent to `https://graph.microsoft.com`. Logs go to stderr and never contain tokens.

## 1. Choose how to sign in

At login you pick one of two options:

| | 1. Own app registration | 2. Microsoft-published app |
|---|---|---|
| What it is | Your own Entra app (public client, no secret) | Microsoft Graph Command Line Tools (`14d82eec-204b-4c2f-b7e8-296a70dab67e`) |
| Setup | One-time registration (below) | None |
| Tenant | Required (tenant ID or domain) | Optional: Enter = any work account (`organizations`), or pin your domain |
| Audit / Conditional Access | Shows under its own name and can be targeted on its own | Shared with Graph PowerShell, so you can't tell the two apart |
| Recommended for | Production / MSP use | Quick personal testing |

With option 2, pin your own domain if you also hold accounts in customer tenants. That stops you signing in to a customer tenant by mistake.

### Registering your own app (option 1 only, about 5 min)

1. Entra admin center → **Identity → Applications → App registrations → New registration**
   - Name: `Claude To Do Planner MCP`
   - Supported account types: **Single tenant**
   - Redirect URI: platform **Public client/native (mobile & desktop)**, value `http://localhost`
2. **API permissions → Add a permission → Microsoft Graph → Delegated**:
   - `User.Read`, `Tasks.ReadWrite` (required)
   - Optional: `GroupMember.Read.All` (lets `planner_list_my_groups` list groups for new plans – admin consent needed)
   - Optional: `User.ReadBasic.All` (lets you assign Planner tasks to people by email)
3. Click **Grant admin consent** (recommended for a managed tenant so users aren't prompted).
4. Do **not** create a client secret. Leave "Allow public client flows" = **No**.
5. Copy the **Application (client) ID** and **Directory (tenant) ID** (or use your verified domain).

## 2. Install

Requires [uv](https://docs.astral.sh/uv/) (it installs Python for you).

```bash
cd C:\Github\ToDo-MCP
uv sync
```

## 3. Sign in once

```bash
uv run todo-planner-mcp login
```

```text
How do you want to sign in to Microsoft 365?
  1. Own app registration (your Entra client ID)
  2. Microsoft-published app (Microsoft Graph Command Line Tools, no registration)
Choose 1 or 2: 1
Application (client) ID: <guid>
Tenant ID or verified domain (e.g. contoso.com.au): contoso.com.au
Time zone ('auto' = detect from Microsoft 365, or e.g. Australia/Sydney) [auto]:
```

Your browser opens for sign-in and MFA. The choice is saved, without secrets, to `%USERPROFILE%\.todo-planner-mcp\settings.json`. Tokens are kept in a separate DPAPI-encrypted cache and refresh silently. Run `login` again to switch options; switching app or tenant clears the old tokens.

To skip the menu: `uv run todo-planner-mcp login --app microsoft --tenant contoso.com.au --timezone Australia/Sydney`
(or `--app own --client-id <guid> --tenant <tenant>`). `status` and `logout` are also available.

You can also sign in from inside Claude. Ask it to "log in to To Do". Claude asks which option you want and then calls the `login` tool, which opens your browser.

### Automatic sign-in when the session expires

You normally sign in once. MSAL refreshes tokens silently, and the refresh token lasts up to 90 days while you keep using it. The session can still end early: a Conditional Access sign-in frequency policy, a password change, revoked sessions or an admin action will all do it.

When that happens, the next request **opens the Microsoft sign-in page in your browser automatically**, pre-filled with your account. Complete sign-in and MFA, and the request carries on.

- Only one sign-in window opens, even if Claude is running several requests at once.
- If you close the window without signing in, it won't reopen for 2 minutes. Ask Claude to "log in to To Do" when you're ready.
- Status checks (`auth_status`, `status`) never open the browser.
- Turn it off or on with `uv run todo-planner-mcp auto-login off|on`, ask Claude to "turn off automatic sign-in", or set `TODO_MCP_AUTO_LOGIN=off`.
- If the window opens behind other windows, look for the browser flashing on the taskbar.

### Time zone detection

Dates you type ("due Friday", "2026-11-01") are interpreted in your time zone. With the default `auto` setting, the server uses, in order:

1. **Your Microsoft 365 time zone**: the one in Outlook → Settings → Time zone, which To Do and Outlook also use. This needs the read-only `MailboxSettings.Read` permission. With option 2 it is requested at your next interactive sign-in. With option 1, add it to the app registration and to `TODO_MCP_EXTRA_SCOPES`. Until it's granted, nothing pops up; the server just uses the next source.
2. **This computer's Windows time zone.**
3. **UTC.**

Microsoft 365 reports Windows zone names (e.g. *AUS Eastern Standard Time*), which the server maps to IANA names (*Australia/Sydney*). To pin a zone instead, run `login` and enter it, or set `TODO_MCP_TIMEZONE`. Ask Claude "check my To Do sign-in status" to see which zone is in use and where it came from.

## 4. Add to Claude Desktop (Chat, Projects and Cowork)

1. **Fully quit Claude Desktop** (tray icon → right-click → Quit). Closing the window isn't enough. While the app is running it rewrites its config from memory and drops entries added in the meantime.
2. From a normal terminal, run:

   ```bash
   uv run --directory C:\Github\ToDo-MCP python scripts\register_claude_desktop.py
   ```

   The script finds the right config file, makes a timestamped backup and adds the entry below. The Store (MSIX) build of Claude Desktop keeps its config in `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\`, not in `%APPDATA%\Claude`. To remove the entry later, add `--remove`.
3. Start Claude Desktop. The tools appear under the connectors/tools menu in any chat, Project or Cowork task.

The entry needs no IDs, because the server reads the saved login choice (see [`claude_desktop_config.example.json`](claude_desktop_config.example.json)):

```json
"mcpServers": {
  "todo-planner": {
    "command": "C:\\Users\\<you>\\.local\\bin\\uv.exe",
    "args": ["run", "--directory", "C:\\Github\\ToDo-MCP", "todo-planner-mcp"]
  }
}
```

For **Claude Code** instead:

```bash
claude mcp add todo-planner --scope user -- uv run --directory C:\Github\ToDo-MCP todo-planner-mcp
```

## Configuration

The saved login choice is used by default. Environment variables override it, for example to enforce one app across a managed fleet. If `TODO_MCP_CLIENT_ID` is set, you can't choose the app at login.

| Variable | Description |
|---|---|
| `TODO_MCP_CLIENT_ID` | Force a client ID (disables the login menu choice) |
| `TODO_MCP_TENANT_ID` | Force a tenant ID, domain or `organizations` |
| `TODO_MCP_TIMEZONE` | `auto` (default: Microsoft 365 → Windows → UTC) or an IANA zone such as `Australia/Sydney` |
| `TODO_MCP_AUTO_LOGIN` | `on` (default) / `off`: open browser sign-in automatically when the session expires |
| `TODO_MCP_EXTRA_SCOPES` | Space-separated extra delegated scopes, e.g. `GroupMember.Read.All User.ReadBasic.All`. Run `login` again after changing |
| `TODO_MCP_CACHE_PATH` | Override token cache location |
| `TODO_MCP_DATA_DIR` | Override the settings + cache folder (default `%USERPROFILE%\.todo-planner-mcp`; avoid AppData, which Claude Desktop redirects) |
| `TODO_MCP_LOG_LEVEL` | `INFO` (default), `DEBUG`, `WARNING` |

## Example prompts

Write naturally; you don't need tool names. Mentioning "To Do" or "Planner" helps Claude pick the right one. These examples are also in the server's description, so Claude sees them in every chat, Project and Cowork task.

**Microsoft To Do (lists = projects)**

| You say | What happens |
|---|---|
| "Show my To Do lists" | `todo_list_lists` |
| "Create a To Do list called *Client Onboarding – Acme*" | `todo_create_list` |
| "Add *Renew SSL certificate* to my Tasks list, due Friday, high importance" | `todo_create_task` |
| "What's still open in my *Work* list?" / "…including completed" | `todo_list_tasks` |
| "Add a reminder tomorrow 9am and a note 'check expiry on the firewall too'" | `todo_update_task` |
| "Mark *Renew SSL certificate* as completed" / "reopen it" | `todo_update_task` |
| "Add checklist steps: backup, renew, verify" / "tick off backup" | `todo_add_checklist_item` / `todo_update_checklist_item` |
| "Move it to the *Admin* list" | `todo_move_task` |
| "Rename the *Work* list to *Work – 2026*" | `todo_rename_list` |
| "Delete the *Old ideas* list" (Claude confirms first) | `todo_delete_list` |

**Microsoft Planner (plans = projects)**

| You say | What happens |
|---|---|
| "Show my Planner tasks" | `planner_my_tasks` |
| "List my plans" / "Which groups can I create a plan in?" | `planner_list_plans` / `planner_list_my_groups` |
| "Create a plan *M365 Migration – Contoso* in the *IT Projects* group with buckets Discovery, Build, Cutover, Hypercare" | `planner_create_plan` |
| "Show open tasks in *M365 Migration – Contoso*, Build bucket only" | `planner_list_tasks` |
| "Add *Licensing audit* to *M365 Migration – Contoso* in Discovery, assign to me, due 2026-11-01, priority important, label *Licensing*" | `planner_create_task` |
| "Move *Licensing audit* to Build, set it in progress and add checklist: export licences, compare, report" | `planner_update_task` |
| "Assign it to jane@contoso.com" (needs `User.ReadBasic.All`) | `planner_update_task` |
| "Rename the *Cutover* bucket to *Go-live*" | `planner_rename_bucket` |
| "Delete the task *Old test*" (Claude confirms first) | `planner_delete_task` |

**Account**

| You say | What happens |
|---|---|
| "Check my To Do sign-in status" (shows account, time zone, auto sign-in) | `auth_status` |
| "Log in to To Do" / "Switch To Do sign-in to my own app registration" | `login` |
| "Turn off automatic sign-in for To Do" | `auto_login` |

**Combined:** "Read my meeting notes below and create Planner tasks in *M365 Migration – Contoso* for every action item, assigned to me, due next week."

## Validation checklist

1. `uv run todo-planner-mcp status` → shows your account, `token valid`.
2. In Claude: "check To Do auth status" → `tokenValid: true`.
3. Create → update → delete a test task in a scratch list; confirm each in the To Do app.
4. Same for a Planner task in a test plan.

## Rollback

- Remove the `todo-planner` entry from `claude_desktop_config.json` and restart Claude.
- `uv run todo-planner-mcp logout`, then delete `%USERPROFILE%\.todo-planner-mcp` (saved choice and token cache).
- Delete the app registration in Entra ID to revoke all access; optionally revoke sessions for the user.

## Troubleshooting

| Symptom | Fix |
|---|---|
| "Not signed in…" / "Sign-in method not chosen" | Run `uv run todo-planner-mcp login` or ask Claude to log in |
| "…opened recently but not completed" | The sign-in window was closed. Ask Claude to "log in to To Do" |
| Dates one day off | Check the zone with "check my To Do sign-in status"; pin it with `TODO_MCP_TIMEZONE` if needed |
| `AADSTS50194` | Your own app is single-tenant: enter your tenant ID/domain instead of `organizations` |
| Option 2 blocked (`AADSTS50105` / `AADSTS65001`) | Your tenant restricts or hasn't consented to Microsoft Graph Command Line Tools: use option 1 |
| `AADSTS65001` consent | Grant admin consent on the app registration |
| `AADSTS50011` redirect mismatch | Ensure redirect URI `http://localhost` under **Mobile and desktop** platform |
| Graph 403 on groups / user lookup | Add the optional scope to `TODO_MCP_EXTRA_SCOPES`, consent, re-login |
| Graph 412 on Planner | Item changed concurrently; just retry |
| Server not visible in Claude | Check `%APPDATA%\Claude\logs\mcp-server-todo-planner.log` |
