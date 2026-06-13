#!/usr/bin/env python3
"""Plugin marketplace CLI for installing agent skills from GitHub repositories."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

SKILLS_DIR = Path(__file__).parent / "skills"
MANIFEST_PATH = Path(__file__).parent / "skills.json"
CACHE_BASE = Path.home() / ".claude" / "plugins" / "cache"

GITHUB_RAW = "https://raw.githubusercontent.com"
GITHUB_API = "https://api.github.com"


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "plugin-marketplace/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"HTTP {e.code} fetching {url}") from e


def parse_frontmatter(text: str) -> dict[str, Any]:
    """Parse YAML-like frontmatter from a SKILL.md file."""
    if not text.startswith("---"):
        return {}
    end = text.find("\n---", 3)
    if end == -1:
        return {}
    block = text[4:end]
    result: dict[str, Any] = {}
    # Handle simple scalar fields only (name, version, description, etc.)
    for line in block.splitlines():
        m = re.match(r'^(\w[\w-]*):\s*"?([^"#\n]*)"?\s*$', line)
        if m:
            result[m.group(1)] = m.group(2).strip()
    return result


def load_manifest() -> dict[str, Any]:
    if MANIFEST_PATH.exists():
        return json.loads(MANIFEST_PATH.read_text())
    return {"skills": {}}


def save_manifest(manifest: dict[str, Any]) -> None:
    MANIFEST_PATH.write_text(json.dumps(manifest, indent=2) + "\n")


def find_skills_in_repo(owner: str, repo: str, branch: str = "main") -> list[str]:
    """Return skill names found under skills/ in the repo by probing common paths."""
    # Try to find a SKILL.md at the standard location skills/<name>/SKILL.md
    # First probe the skills/ directory listing is not possible without API auth,
    # so we derive the skill name from the repo name convention: <name>-skill
    if repo.endswith("-skill"):
        return [repo[: -len("-skill")]]
    return [repo]


def install_skill_files(
    owner: str, repo: str, skill_name: str, branch: str, version: str
) -> Path:
    """Download skill files into the project skills dir and the cache."""
    skill_dir = SKILLS_DIR / skill_name
    cache_dir = CACHE_BASE / repo / skill_name / version / "skills" / skill_name

    skill_dir.mkdir(parents=True, exist_ok=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    base = f"{GITHUB_RAW}/{owner}/{repo}/{branch}/skills/{skill_name}"
    cache_base_url = base  # same source

    # Known files to download (SKILL.md + all scripts)
    script_files = [
        "scripts/last30days.py",
        "scripts/store.py",
        "scripts/build-skill.sh",
        "scripts/compare.sh",
        "scripts/setup-keychain.sh",
        "scripts/lib/__init__.py",
        "scripts/lib/dates.py",
        "scripts/lib/digg.py",
        "scripts/lib/env.py",
        "scripts/lib/http.py",
        "scripts/lib/log.py",
        "scripts/lib/ui.py",
    ]

    downloaded = []

    # SKILL.md first
    for dest_root in (skill_dir, cache_dir):
        url = f"{base}/SKILL.md"
        try:
            content = fetch(url)
            (dest_root / "SKILL.md").write_bytes(content)
            if dest_root == skill_dir:
                downloaded.append("SKILL.md")
        except RuntimeError as e:
            print(f"  warning: {e}", file=sys.stderr)

    # Scripts
    for rel in script_files:
        url = f"{base}/{rel}"
        try:
            content = fetch(url)
        except RuntimeError:
            continue  # optional files — skip if missing
        for dest_root in (skill_dir, cache_dir):
            dest = dest_root / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(content)
        downloaded.append(rel)

    return skill_dir


def cmd_add(args: argparse.Namespace) -> int:
    ref = args.ref  # e.g. "mvanhorn/last30days-skill"
    parts = ref.split("/")
    if len(parts) != 2:
        print(f"error: expected <owner>/<repo>, got: {ref}", file=sys.stderr)
        return 1

    owner, repo = parts
    branch = getattr(args, "branch", "main") or "main"

    print(f"Fetching skill metadata from {owner}/{repo}...")

    skill_names = find_skills_in_repo(owner, repo, branch)
    skill_name = skill_names[0]

    # Fetch SKILL.md to read version and metadata
    skill_md_url = f"{GITHUB_RAW}/{owner}/{repo}/{branch}/skills/{skill_name}/SKILL.md"
    try:
        skill_md = fetch(skill_md_url).decode()
    except RuntimeError as e:
        print(f"error: could not fetch SKILL.md: {e}", file=sys.stderr)
        return 1

    meta = parse_frontmatter(skill_md)
    version = meta.get("version", "0.0.0")
    description = meta.get("description", "")

    print(f"Installing {skill_name} v{version}...")
    skill_dir = install_skill_files(owner, repo, skill_name, branch, version)

    manifest = load_manifest()
    manifest["skills"][skill_name] = {
        "version": version,
        "description": description,
        "source": f"{owner}/{repo}",
        "branch": branch,
        "path": str(skill_dir.relative_to(Path(__file__).parent)),
        "homepage": meta.get("homepage", f"https://github.com/{owner}/{repo}"),
        "installed_at": _now_iso(),
    }
    save_manifest(manifest)

    print(f"\nInstalled skill: /{skill_name}")
    print(f"  version  : {version}")
    print(f"  source   : {owner}/{repo}")
    print(f"  path     : {skill_dir}")
    if description:
        print(f"  about    : {description[:80]}...")
    print(f"\nUse it: /{skill_name} <topic>")
    return 0


def cmd_list(args: argparse.Namespace) -> int:
    manifest = load_manifest()
    skills = manifest.get("skills", {})
    if not skills:
        print("No skills installed. Use: plugin marketplace add <owner>/<repo>")
        return 0
    print(f"{'NAME':<20} {'VERSION':<10} {'SOURCE'}")
    print("-" * 60)
    for name, info in sorted(skills.items()):
        print(f"{name:<20} {info.get('version', '?'):<10} {info.get('source', '?')}")
    return 0


def cmd_remove(args: argparse.Namespace) -> int:
    name = args.name
    manifest = load_manifest()
    if name not in manifest.get("skills", {}):
        print(f"error: skill '{name}' is not installed", file=sys.stderr)
        return 1
    skill_dir = SKILLS_DIR / name
    if skill_dir.exists():
        shutil.rmtree(skill_dir)
    del manifest["skills"][name]
    save_manifest(manifest)
    print(f"Removed skill: {name}")
    return 0


def _now_iso() -> str:
    import datetime
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="plugin",
        description="Plugin marketplace for agent skills",
    )
    sub = p.add_subparsers(dest="command", required=True)

    mp = sub.add_parser("marketplace", help="Marketplace operations")
    mp_sub = mp.add_subparsers(dest="subcommand", required=True)

    add_p = mp_sub.add_parser("add", help="Install a skill from GitHub")
    add_p.add_argument("ref", metavar="owner/repo", help="GitHub repo (e.g. mvanhorn/last30days-skill)")
    add_p.add_argument("--branch", default="main", help="Branch to install from (default: main)")

    mp_sub.add_parser("list", help="List installed skills")

    rm_p = mp_sub.add_parser("remove", help="Remove an installed skill")
    rm_p.add_argument("name", help="Skill name to remove")

    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "marketplace":
        if args.subcommand == "add":
            return cmd_add(args)
        elif args.subcommand == "list":
            return cmd_list(args)
        elif args.subcommand == "remove":
            return cmd_remove(args)

    parser.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
