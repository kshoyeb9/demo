import pytest

from swipebot.config import Preferences
from swipebot.profile import parse_profile
from swipebot.scorer import score_profile

CARD = """Anna 29
Recently Active
5 miles away
Love hiking and good books. Dog mum.
Like
Nope"""


def test_parse_profile():
    p = parse_profile(CARD)
    assert (p.name, p.age) == ("Anna", 29)
    assert p.distance_km == pytest.approx(8.05, abs=0.01)
    assert p.bio == "Love hiking and good books. Dog mum."


def test_parse_km_and_comma_name():
    p = parse_profile("Jean-Luc, 34\n12 km away\nChef")
    assert (p.name, p.age, p.distance_km, p.bio) == ("Jean-Luc", 34, 12, "Chef")


def prefs(**kw):
    base = dict(min_age=25, max_age=35, max_distance_km=25,
                dealbreaker_keywords=["smoker"], keyword_points={"hiking": 2, "books": 1},
                base_score=0, like_threshold=2)
    return Preferences(**{**base, **kw})


def test_likes_on_keywords():
    d = score_profile(parse_profile(CARD), prefs())
    assert d.like and d.score == 3


def test_hard_filters():
    assert not score_profile(parse_profile("Bo 22\nhiking"), prefs()).like
    assert not score_profile(parse_profile("Bo 40\nhiking"), prefs()).like
    assert not score_profile(parse_profile("Bo 30\n50 km away\nhiking"), prefs()).like
    assert not score_profile(parse_profile("Bo 30\nhiking, social smoker"), prefs()).like
    assert not score_profile(parse_profile("Bo 30"), prefs(require_bio=True, like_threshold=0)).like


def test_below_threshold_passes():
    assert not score_profile(parse_profile("Bo 30\nbooks"), prefs()).like


def test_unreadable_uses_default():
    assert not score_profile(parse_profile(""), prefs()).like
    assert score_profile(parse_profile(""), prefs(unreadable_decision="like")).like
