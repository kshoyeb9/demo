from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml


@dataclass
class Preferences:
    min_age: int | None = None
    max_age: int | None = None
    max_distance_km: float | None = None
    require_bio: bool = False
    dealbreaker_keywords: list[str] = field(default_factory=list)
    keyword_points: dict[str, float] = field(default_factory=dict)
    base_score: float = 1
    like_threshold: float = 1
    unreadable_decision: str = "pass"


@dataclass
class Limits:
    max_swipes_per_session: int = 60
    max_likes_per_day: int = 40
    like_ratio_cap: float = 0.6


@dataclass
class Pacing:
    view_seconds: tuple[float, float] = (4, 12)
    break_every: tuple[int, int] = (15, 25)
    break_seconds: tuple[float, float] = (45, 150)


@dataclass
class Config:
    preferences: Preferences = field(default_factory=Preferences)
    limits: Limits = field(default_factory=Limits)
    pacing: Pacing = field(default_factory=Pacing)


def load_config(path: str | Path) -> Config:
    raw = yaml.safe_load(Path(path).read_text()) or {}
    prefs = Preferences(**(raw.get("preferences") or {}))
    if prefs.unreadable_decision not in ("like", "pass"):
        raise ValueError("preferences.unreadable_decision must be 'like' or 'pass'")
    pacing_raw = raw.get("pacing") or {}
    pacing = Pacing(**{k: tuple(v) for k, v in pacing_raw.items()})
    return Config(
        preferences=prefs,
        limits=Limits(**(raw.get("limits") or {})),
        pacing=pacing,
    )
