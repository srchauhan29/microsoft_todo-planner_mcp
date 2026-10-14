"""Add (or update) the todo-planner MCP server in Claude Desktop's config.

Run this with Claude Desktop fully QUIT (tray icon -> Quit). The running app keeps its
settings in memory and rewrites the config file, which drops entries added while it runs.

On Windows the Store/MSIX build of Claude Desktop keeps its config in a private package
folder (%LOCALAPPDATA%\\Packages\\Claude_*\\LocalCache\\Roaming\\Claude); the classic
installer uses %APPDATA%\\Claude. Every config file found is updated, each with a backup.

Usage:  uv run python scripts\\register_claude_desktop.py [--remove]
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

SERVER_NAME = "todo-planner"
REPO = Path(__file__).resolve().parent.parent


def claude_running() -> bool:
    """True if Claude Desktop (not the Claude Code CLI, also named claude.exe) is running."""
    if os.name != "nt":
        return False
    query = (
        "Get-Process -Name claude -ErrorAction SilentlyContinue | Where-Object { -not $_.Path -or "
        "$_.Path -match '\\\\WindowsApps\\\\Claude_|\\\\AnthropicClaude\\\\' } | Measure-Object | "
        "Select-Object -ExpandProperty Count"
    )
    out = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", query], capture_output=True, text=True
    ).stdout.strip()
    return out not in ("", "0")


def config_paths() -> list[Path]:
    local, roaming = Path(os.environ.get("LOCALAPPDATA", "")), Path(os.environ.get("APPDATA", ""))
    found = sorted(local.glob("Packages/Claude_*/LocalCache/Roaming/Claude/claude_desktop_config.json"))
    classic = roaming / "Claude" / "claude_desktop_config.json"
    if classic.exists():
        found.append(classic)
    if not found:
        pkg_dirs = sorted(local.glob("Packages/Claude_*/LocalCache/Roaming/Claude"))
        found = [(pkg_dirs[0] if pkg_dirs else classic.parent) / "claude_desktop_config.json"]
    return found


def server_entry() -> dict:
    uv = shutil.which("uv") or str(Path.home() / ".local" / "bin" / "uv.exe")
    return {"command": uv, "args": ["run", "--directory", str(REPO), "todo-planner-mcp"]}


def main() -> int:
    remove = "--remove" in sys.argv
    if claude_running():
        print("Claude Desktop is still running. Quit it from the tray icon (right-click -> Quit), then re-run.")
        return 1

    for path in config_paths():
        cfg: dict = {}
        if path.exists():
            cfg = json.loads(path.read_text(encoding="utf-8"))
            backup = path.with_name(f"{path.name}.bak-{datetime.now():%Y%m%d-%H%M%S}")
            shutil.copy2(path, backup)
            print(f"Backup: {backup}")
        servers = cfg.setdefault("mcpServers", {})
        if remove:
            servers.pop(SERVER_NAME, None)
        else:
            servers[SERVER_NAME] = server_entry()
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
        json.loads(tmp.read_text(encoding="utf-8"))  # validate before replacing
        os.replace(tmp, path)
        print(f"{'Removed from' if remove else 'Updated'}: {path}")

    print("Done. Start Claude Desktop again; the todo-planner tools appear in chats, Projects and Cowork.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
