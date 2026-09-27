from __future__ import annotations

from dataclasses import dataclass, field

from swipebot.config import Preferences
from swipebot.profile import Profile


@dataclass
class Decision:
    like: bool
    score: float
    reasons: list[str] = field(default_factory=list)


def score_profile(profile: Profile, prefs: Preferences) -> Decision:
    if not profile.readable:
        return Decision(prefs.unreadable_decision == "like", 0, ["unreadable profile"])

    if profile.age is not None:
        if prefs.min_age is not None and profile.age < prefs.min_age:
            return Decision(False, 0, [f"age {profile.age} < {prefs.min_age}"])
        if prefs.max_age is not None and profile.age > prefs.max_age:
            return Decision(False, 0, [f"age {profile.age} > {prefs.max_age}"])

    if (
        prefs.max_distance_km is not None
        and profile.distance_km is not None
        and profile.distance_km > prefs.max_distance_km
    ):
        return Decision(False, 0, [f"distance {profile.distance_km:.0f}km > {prefs.max_distance_km}km"])

    bio = profile.bio.lower()
    if prefs.require_bio and not bio:
        return Decision(False, 0, ["empty bio"])

    for word in prefs.dealbreaker_keywords:
        if word.lower() in bio:
            return Decision(False, 0, [f"dealbreaker '{word}'"])

    score = prefs.base_score
    reasons = []
    for word, points in prefs.keyword_points.items():
        if word.lower() in bio:
            score += points
            reasons.append(f"+{points:g} '{word}'")

    like = score >= prefs.like_threshold
    reasons.append(f"score {score:g} {'>=' if like else '<'} {prefs.like_threshold:g}")
    return Decision(like, score, reasons)
