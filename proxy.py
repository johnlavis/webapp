#!/usr/bin/env python3
"""Headroom proxy manager — install and start the headroom compression proxy."""

from __future__ import annotations

import argparse
import subprocess
import sys


DEFAULT_PORT = 8787
PACKAGE = "headroom-ai[proxy]"


def _headroom_installed() -> bool:
    try:
        import importlib.util
        return importlib.util.find_spec("headroom") is not None
    except Exception:
        return False


def _install() -> None:
    print(f"Installing {PACKAGE}...")
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", PACKAGE],
        stdout=subprocess.DEVNULL,
    )
    print("Done.\n")


def cmd_start(args: argparse.Namespace) -> int:
    port: int = args.port

    if not _headroom_installed():
        _install()

    print(f"Starting headroom proxy on port {port}...\n")
    print(f"  Claude Code  :  ANTHROPIC_BASE_URL=http://localhost:{port} claude")
    print(f"  OpenAI compat:  OPENAI_BASE_URL=http://localhost:{port}/v1 your-app\n")

    try:
        subprocess.run(["headroom", "proxy", "--port", str(port)], check=True)
    except FileNotFoundError:
        # headroom CLI not on PATH yet — try via module
        subprocess.run(
            [sys.executable, "-m", "headroom", "proxy", "--port", str(port)],
            check=True,
        )
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="proxy",
        description="Headroom proxy manager",
    )
    sub = p.add_subparsers(dest="command", required=True)

    start_p = sub.add_parser("start", help="Install (if needed) and start the proxy")
    start_p.add_argument(
        "--port",
        type=int,
        default=DEFAULT_PORT,
        help=f"Port to listen on (default: {DEFAULT_PORT})",
    )

    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "start":
        return cmd_start(args)

    parser.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
