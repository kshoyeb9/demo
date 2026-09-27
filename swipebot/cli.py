from __future__ import annotations

import argparse
from pathlib import Path

from playwright.sync_api import sync_playwright

from swipebot.config import load_config
from swipebot.platforms import PLATFORMS
from swipebot.runner import DATA_DIR, run
from swipebot.scorer import Decision


def _launch(p, platform_name: str, headless: bool):
    # A persistent profile keeps you logged in between runs, so you only log
    # in by hand once and the bot never sees your password.
    profile_dir = DATA_DIR / "browser" / platform_name
    profile_dir.mkdir(parents=True, exist_ok=True)
    context = p.chromium.launch_persistent_context(
        str(profile_dir), headless=headless, viewport={"width": 1280, "height": 900}
    )
    page = context.pages[0] if context.pages else context.new_page()
    return context, page


def _ask(text: str, decision: Decision) -> bool | None:
    print("\n" + "-" * 60 + f"\n{text[:600]}\n" + "-" * 60)
    suggestion = "LIKE" if decision.like else "PASS"
    print(f"Suggested: {suggestion} ({'; '.join(decision.reasons)})")
    answer = input("[Enter]=accept  l=like  p=pass  q=quit > ").strip().lower()
    if answer == "q":
        return None
    if answer in ("l", "p"):
        return answer == "l"
    return decision.like


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(prog="swipebot")
    sub = parser.add_subparsers(dest="command", required=True)

    login = sub.add_parser("login", help="open a browser to log in by hand (once)")
    login.add_argument("platform", choices=PLATFORMS)

    go = sub.add_parser("run", help="start swiping")
    go.add_argument("platform", choices=PLATFORMS)
    go.add_argument("--config", default="config.yaml")
    go.add_argument("--confirm", action="store_true",
                    help="show each suggestion and wait for you to accept or override it")
    go.add_argument("--headless", action="store_true",
                    help="hide the browser (more likely to be detected)")
    go.add_argument("--url", help="override the start URL (for testing)")

    args = parser.parse_args(argv)
    cls = PLATFORMS[args.platform]

    with sync_playwright() as p:
        if args.command == "login":
            context, page = _launch(p, args.platform, headless=False)
            page.goto(cls.url)
            input(f"Log in to {args.platform} in the browser window, then press Enter here...")
            context.close()
            print("Session saved.")
            return

        config = load_config(Path(args.config))
        context, page = _launch(p, args.platform, headless=args.headless)
        try:
            platform = cls(page)
            platform.open(args.url)
            page.wait_for_timeout(4_000)
            run(platform, config, confirm=_ask if args.confirm else None)
        finally:
            context.close()
