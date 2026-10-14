"""Build the Claude Desktop extension: dist/todo-planner-<version>.mcpb

- Generates manifest.json (MCPB v0.4, uv runtime) from pyproject.toml and the live tool list,
  so the manifest can't drift from the code.
- Optionally validates it against the official JSON schema (--schema path/to/schema.json).
- Packs the bundle (a zip) from an explicit allow-list of files - never .venv, tokens or settings.

Usage:  uv run python scripts\\build_mcpb.py [--schema mcpb-manifest-v0.4.schema.json]
"""

from __future__ import annotations

import argparse
import asyncio
import json
import struct
import sys
import tomllib
import zipfile
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from todo_planner_mcp.app import EXAMPLE_PROMPTS, mcp  # noqa: E402
from todo_planner_mcp import server  # noqa: E402,F401  (registers account tools)

BUNDLE_FILES = ["manifest.json", "icon.png", "pyproject.toml", "uv.lock", "README.md", "LICENSE", "run_server.py"]
BUNDLE_DIRS = ["src/todo_planner_mcp"]


def make_icon(path: Path, size: int = 256) -> None:
    """Draw a simple rounded blue tile with a white tick (no imaging libraries needed)."""
    bg, fg = (0x1F, 0x6F, 0xEB), (0xFF, 0xFF, 0xFF)
    r = size * 0.18
    tick = [(0.25, 0.53), (0.43, 0.70), (0.76, 0.34)]  # polyline in unit coords
    width = size * 0.075

    def seg_dist(px, py, a, b):
        ax, ay, bx, by = a[0] * size, a[1] * size, b[0] * size, b[1] * size
        dx, dy = bx - ax, by - ay
        t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
        return ((px - ax - t * dx) ** 2 + (py - ay - t * dy) ** 2) ** 0.5

    rows = []
    for y in range(size):
        row = bytearray([0])
        for x in range(size):
            cx = min(max(x, r), size - 1 - r)
            cy = min(max(y, r), size - 1 - r)
            inside = (x - cx) ** 2 + (y - cy) ** 2 <= r * r
            if not inside:
                row += bytes((0, 0, 0, 0))
                continue
            on_tick = min(seg_dist(x, y, tick[0], tick[1]), seg_dist(x, y, tick[1], tick[2])) <= width
            row += bytes((*(fg if on_tick else bg), 255))
        rows.append(bytes(row))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(b"".join(rows), 9)) + chunk(b"IEND", b"")
    path.write_bytes(png)


async def tool_list() -> list[dict]:
    tools = await mcp.list_tools()
    return [{"name": t.name, "description": (t.description or "").strip().splitlines()[0]} for t in tools]


def build_manifest(version: str, tools: list[dict]) -> dict:
    prompts_md = "\n".join(f"- {line[2:]}" if line.startswith("- ") else line for line in EXAMPLE_PROMPTS.splitlines())
    return {
        "manifest_version": "0.4",
        "name": "todo-planner",
        "display_name": "Microsoft To Do & Planner",
        "version": version,
        "description": "Create, update and delete Microsoft To Do and Microsoft Planner tasks, lists and plans from Claude.",
        "long_description": (
            "Manage your **Microsoft To Do** lists and tasks and **Microsoft Planner** plans, buckets and tasks "
            "through Microsoft Graph, as the signed-in Microsoft 365 user.\n\n"
            "**Sign-in:** after installing, ask Claude *\"Log in to To Do\"* and choose either your organisation's "
            "own Entra app registration or the Microsoft-published *Microsoft Graph Command Line Tools* app. "
            "Your browser opens for Microsoft sign-in and MFA. No passwords or secrets are stored; tokens are "
            "cached per user, encrypted with Windows DPAPI. If the session expires, the sign-in window opens "
            "automatically.\n\n"
            "**Time zone** is detected from your Microsoft 365 settings (or this computer).\n\n"
            "**Requires** [uv](https://docs.astral.sh/uv/) on this computer; it installs Python and the "
            "dependencies automatically on first start.\n\n" + prompts_md
        ),
        "author": {"name": "Augmented Systems"},
        "icon": "icon.png",
        "server": {
            "type": "uv",
            "entry_point": "run_server.py",
            "mcp_config": {
                "command": "uv",
                "args": ["run", "--directory", "${__dirname}", "--frozen", "--no-dev", "run_server.py"],
                "env": {
                    # Keep uv-managed Python outside AppData: Claude Desktop (MSIX) redirects AppData writes.
                    "UV_PYTHON_INSTALL_DIR": "${HOME}/.local/share/uv/python",
                    "TODO_MCP_CLIENT_ID": "${user_config.client_id}",
                    "TODO_MCP_TENANT_ID": "${user_config.tenant}",
                    "TODO_MCP_TIMEZONE": "${user_config.timezone}",
                    "TODO_MCP_AUTO_LOGIN": "${user_config.auto_login}",
                    "TODO_MCP_EXTRA_SCOPES": "${user_config.extra_scopes}",
                },
            },
        },
        "tools": tools,
        "keywords": ["microsoft-to-do", "planner", "microsoft-365", "tasks", "graph"],
        "license": "MIT",
        "repository": {"type": "git", "url": "https://github.com/srchauhan29/microsoft_todo-planner_mcp"},
        "privacy_policies": ["https://privacy.microsoft.com/privacystatement"],
        "compatibility": {"platforms": ["win32", "darwin"], "runtimes": {"python": ">=3.11"}},
        "user_config": {
            "client_id": {
                "type": "string",
                "title": "Own app registration client ID (optional)",
                "description": "Leave blank to choose at sign-in. Set it to force your organisation's own Entra app "
                "(option 1) for everyone using this install. Not a secret.",
                "required": False,
                "default": "",
            },
            "tenant": {
                "type": "string",
                "title": "Tenant ID or domain (optional)",
                "description": "Restricts sign-in to your tenant, e.g. contoso.com.au. Recommended if you also have "
                "accounts in customer tenants. Leave blank to use the value chosen at sign-in.",
                "required": False,
                "default": "",
            },
            "timezone": {
                "type": "string",
                "title": "Time zone",
                "description": "'auto' detects it from Microsoft 365 (then Windows). Or an IANA zone such as Australia/Sydney.",
                "required": False,
                "default": "auto",
            },
            "auto_login": {
                "type": "boolean",
                "title": "Open sign-in automatically when the session expires",
                "description": "Opens the Microsoft sign-in page in your browser when needed; the request then continues.",
                "required": False,
                "default": True,
            },
            "extra_scopes": {
                "type": "string",
                "title": "Extra Graph permissions (optional)",
                "description": "Space-separated, e.g. 'GroupMember.Read.All User.ReadBasic.All' to list groups and "
                "assign tasks by email. Sign in again after changing.",
                "required": False,
                "default": "",
            },
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--schema", type=Path, help="MCPB manifest JSON schema to validate against")
    args = parser.parse_args()

    version = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    make_icon(ROOT / "icon.png")
    manifest = build_manifest(version, asyncio.run(tool_list()))
    (ROOT / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    print(f"manifest.json: v{version}, {len(manifest['tools'])} tools")

    if args.schema:
        import jsonschema

        jsonschema.validate(manifest, json.loads(args.schema.read_text(encoding="utf-8")))
        print(f"Schema validation passed ({args.schema.name})")

    if not (ROOT / "uv.lock").exists():
        raise SystemExit("uv.lock missing - run `uv lock` first (the bundle runs with --frozen).")

    out = ROOT / "dist" / f"todo-planner-{version}.mcpb"
    out.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in BUNDLE_FILES:
            zf.write(ROOT / name, name)
        for folder in BUNDLE_DIRS:
            for f in sorted((ROOT / folder).rglob("*")):
                if f.is_file() and "__pycache__" not in f.parts and f.suffix != ".pyc":
                    zf.write(f, f.relative_to(ROOT).as_posix())
    with zipfile.ZipFile(out) as zf:
        names = zf.namelist()
    print(f"Built {out} ({out.stat().st_size // 1024} KB, {len(names)} files)")
    for n in names:
        print("  ", n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
