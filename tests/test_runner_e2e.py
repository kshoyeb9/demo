"""Drives a mock swipe page in real Chromium to check the full loop."""
import json
import os

import pytest
from playwright.sync_api import sync_playwright

from swipebot.config import Config, Limits, Pacing, Preferences
from swipebot.platforms import Tinder
from swipebot.runner import run

PROFILES = [
    ("Anna 29", "5 km away", "Hiking every weekend"),
    ("Bea 41", "3 km away", "Hiking too"),
    ("Cleo 30", "2 km away", "Just here for fun"),
    ("Dana 27", "80 km away", "Hiking"),
    ("Eve 31", "1 km away", "Books and hiking"),
]

MOCK = """
<main><div id="card" data-testid="swipe-card"></div></main>
<div id="log"></div>
<script>
const profiles = %s;
let i = 0, expanded = false;
const render = () => {
  const card = document.getElementById('card');
  if (i >= profiles.length) { card.remove(); document.querySelector('main').innerText = "There's no one new around you"; return; }
  const [head, dist, bio] = profiles[i];
  // Bio is only visible when the profile is expanded, like Tinder web.
  card.innerText = [head, dist].concat(expanded ? [bio] : []).join('\\n');
};
document.addEventListener('keydown', e => {
  if (e.key === 'ArrowUp') expanded = true;
  else if (e.key === 'ArrowDown') expanded = false;
  else if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
    document.getElementById('log').innerText += (e.key === 'ArrowRight' ? 'L' : 'P');
    i++;
  } else return;
  render();
});
render();
</script>
""" % json.dumps(PROFILES)


@pytest.fixture
def page():
    with sync_playwright() as p:
        # Lets CI point at a preinstalled browser whose version differs from the package.
        browser = p.chromium.launch(executable_path=os.environ.get("SWIPEBOT_CHROMIUM") or None)
        yield browser.new_page()
        browser.close()


def config(**limits):
    return Config(
        preferences=Preferences(min_age=25, max_age=35, max_distance_km=25,
                                keyword_points={"hiking": 1}, base_score=0, like_threshold=1),
        limits=Limits(**{"max_swipes_per_session": 50, "max_likes_per_day": 50,
                         "like_ratio_cap": 1.0, **limits}),
        pacing=Pacing(view_seconds=(0, 0), break_every=(100, 100), break_seconds=(0, 0)),
    )


def test_swipes_according_to_rules(page, tmp_path):
    page.set_content(MOCK)
    stats = run(Tinder(page), config(), data_dir=tmp_path, sleep=lambda s: None, echo=print)
    assert page.inner_text("#log") == "LPPPL"
    assert stats == {"swipes": 5, "likes": 2, "passes": 3}
    assert "Anna" not in (tmp_path / "decisions.csv").read_text()


def test_daily_like_limit_persists(page, tmp_path):
    page.set_content(MOCK)
    run(Tinder(page), config(max_likes_per_day=1), data_dir=tmp_path, sleep=lambda s: None)
    assert page.inner_text("#log") == "LPPP"  # stops before liking Eve
    page.set_content(MOCK)
    run(Tinder(page), config(max_likes_per_day=1), data_dir=tmp_path, sleep=lambda s: None)
    assert page.inner_text("#log") == ""  # already at today's limit


def test_confirm_can_override_and_quit(page, tmp_path):
    page.set_content(MOCK)
    answers = iter([False, True, None])
    run(Tinder(page), config(), confirm=lambda text, d: next(answers),
        data_dir=tmp_path, sleep=lambda s: None)
    assert page.inner_text("#log") == "PL"
