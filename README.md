# swipebot

Auto-swipes on the **Tinder** and **Bumble** web apps using rules you define
(age, distance, bio keywords, dealbreakers).

> ⚠️ Automated swiping breaks both apps' Terms of Service. Your account can be
> shadowbanned or banned. The default limits and human-like pacing lower the
> risk but don't remove it. Use at your own risk.

## Setup

```bash
pip install -r requirements.txt
playwright install chromium
cp config.example.yaml config.yaml   # then edit your preferences
```

Log in once per app. A real browser opens; log in by hand and press Enter in
the terminal. The session is saved in `~/.swipebot/browser/<app>`, so the bot
never handles your password.

```bash
python -m swipebot login tinder
python -m swipebot login bumble
```

## Running

Start with `--confirm`: the bot shows each profile and its suggestion, and you
accept (Enter), override (`l`/`p`), or quit (`q`). Use this to tune your rules
before letting it run on its own.

```bash
python -m swipebot run tinder --confirm
python -m swipebot run bumble             # fully automatic
```

## How it works

- **Reading profiles:** it reads the visible text of the current card and
  parses name, age, distance and bio. On Tinder it presses ↑ to expand the
  full profile first.
- **Deciding:** hard filters (age, distance, empty bio, dealbreaker words)
  mean an automatic pass. Anything else gets `base_score` plus points for each
  keyword in the bio, and is liked if it reaches `like_threshold`.
- **Swiping:** it uses the apps' keyboard shortcuts (→ like, ← pass), which
  change less often than their page markup.
- **Safety limits:** random viewing time per profile, periodic breaks, a
  per-session swipe cap, a daily like cap (tracked across runs), and a
  like-ratio cap. It stops when the page says you're out of likes or profiles,
  and closes match and upsell popups with Escape.
- **Log:** `~/.swipebot/decisions.csv` records each action, score and reason.
  It doesn't store names, bios or photos.

## When it breaks

Both sites change their markup without notice. If the bot says the card
didn't change or every profile is "unreadable", update `card_selectors`,
`stop_patterns` or `popup_patterns` in `swipebot/platforms/tinder.py` or
`bumble.py`.

## Tests

```bash
pytest   # set SWIPEBOT_CHROMIUM=/path/to/chrome if Playwright's own browser isn't installed
```

The end-to-end tests drive a mock swipe page in a real browser, not the live
sites.
