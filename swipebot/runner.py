from __future__ import annotations

import csv
import json
import random
import time
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from swipebot.config import Config
from swipebot.platforms.base import Platform
from swipebot.profile import parse_profile
from swipebot.scorer import Decision, score_profile

DATA_DIR = Path.home() / ".swipebot"


class DailyCounter:
    """Persists likes per platform per day so limits hold across sessions."""

    def __init__(self, path: Path):
        self.path = path
        try:
            self.data = json.loads(path.read_text())
        except (FileNotFoundError, json.JSONDecodeError):
            self.data = {}
        today = date.today().isoformat()
        self.data = {today: self.data.get(today, {})}
        self.today = self.data[today]

    def likes(self, platform: str) -> int:
        return self.today.get(platform, 0)

    def add_like(self, platform: str) -> None:
        self.today[platform] = self.likes(platform) + 1
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.data))


def _log(path: Path, platform: str, decision: Decision, action: str) -> None:
    # Only the decision is logged, never names, bios or photos of other people.
    path.parent.mkdir(parents=True, exist_ok=True)
    new = not path.exists()
    with path.open("a", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["time", "platform", "action", "score", "reasons"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), platform, action,
                    decision.score, "; ".join(decision.reasons)])


def run(
    platform: Platform,
    config: Config,
    *,
    confirm: Callable[[str, Decision], bool | None] | None = None,
    data_dir: Path = DATA_DIR,
    sleep: Callable[[float], None] = time.sleep,
    echo: Callable[[str], None] = print,
) -> dict[str, int]:
    """Swipe until a limit or stop condition is hit. Returns session stats.

    confirm, if given, is called with (profile_text, decision) before each
    swipe and returns True (like), False (pass) or None (stop the session).
    """
    limits, pacing = config.limits, config.pacing
    counter = DailyCounter(data_dir / "daily.json")
    log_path = data_dir / "decisions.csv"
    stats = {"swipes": 0, "likes": 0, "passes": 0}
    next_break = random.randint(*pacing.break_every)

    while stats["swipes"] < limits.max_swipes_per_session:
        platform.dismiss_popups()
        if reason := platform.stop_reason():
            echo(f"Stopping: page says '{reason}'")
            break

        text = platform.read_profile_text()
        profile = parse_profile(text)
        decision = score_profile(profile, config.preferences)
        label = f"{profile.name or '?'}, {profile.age or '?'}"

        sleep(random.uniform(*pacing.view_seconds))

        like = decision.like
        if confirm is not None:
            choice = confirm(text, decision)
            if choice is None:
                echo("Stopped by user.")
                break
            like = choice

        if like and counter.likes(platform.name) >= limits.max_likes_per_day:
            echo(f"Daily like limit ({limits.max_likes_per_day}) reached.")
            break
        ratio_after = (stats["likes"] + 1) / (stats["swipes"] + 1)
        if like and confirm is None and stats["swipes"] >= 5 and ratio_after > limits.like_ratio_cap:
            like = False
            decision.reasons.append("like-ratio cap")

        before = platform.card_text()
        if like:
            platform.like()
            counter.add_like(platform.name)
            stats["likes"] += 1
        else:
            platform.pass_()
            stats["passes"] += 1
        stats["swipes"] += 1
        action = "LIKE" if like else "PASS"
        _log(log_path, platform.name, decision, action)
        echo(f"[{stats['swipes']}] {action:4} {label:20} {'; '.join(decision.reasons)}")

        if not platform.wait_for_next_card(before):
            echo("Card didn't change after swiping; the page layout may have changed. Stopping.")
            break

        if stats["swipes"] >= next_break:
            pause = random.uniform(*pacing.break_seconds)
            echo(f"Taking a {pause:.0f}s break...")
            sleep(pause)
            next_break = stats["swipes"] + random.randint(*pacing.break_every)

    echo(f"Session done: {stats}")
    return stats
