from __future__ import annotations

import re

from playwright.sync_api import Page


class Platform:
    """One dating site's web app.

    Swipes use each site's keyboard shortcuts (Right = like, Left = pass),
    which have been far more stable than their obfuscated CSS classes.
    Card selectors are tried in order, falling back to the whole <main>.
    """

    name: str
    url: str
    card_selectors: list[str] = []
    # Visible text that means we should stop the session.
    stop_patterns: list[str] = []
    # Visible text for overlays (match screens, upsells) that Escape closes.
    popup_patterns: list[str] = []

    def __init__(self, page: Page):
        self.page = page

    def open(self, url: str | None = None) -> None:
        self.page.goto(url or self.url, wait_until="domcontentloaded")

    def card_text(self) -> str:
        for selector in [*self.card_selectors, "main"]:
            loc = self.page.locator(selector).first
            if loc.count() and loc.is_visible():
                text = loc.inner_text(timeout=5_000).strip()
                if text:
                    return text
        return ""

    def read_profile_text(self) -> str:
        """Text used for scoring. Subclasses can expand the profile first."""
        return self.card_text()

    def like(self) -> None:
        self.page.keyboard.press("ArrowRight")

    def pass_(self) -> None:
        self.page.keyboard.press("ArrowLeft")

    def _page_text(self) -> str:
        return self.page.locator("body").inner_text(timeout=5_000)

    def stop_reason(self) -> str | None:
        text = self._page_text()
        for pattern in self.stop_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                return pattern
        return None

    def dismiss_popups(self) -> bool:
        text = self._page_text()
        if any(re.search(p, text, re.IGNORECASE) for p in self.popup_patterns):
            self.page.keyboard.press("Escape")
            self.page.wait_for_timeout(800)
            return True
        return False

    def wait_for_next_card(self, previous_text: str, timeout_ms: int = 10_000) -> bool:
        """Wait until the card on screen differs from the one we just swiped."""
        waited = 0
        while waited < timeout_ms:
            self.page.wait_for_timeout(250)
            waited += 250
            self.dismiss_popups()
            if self.card_text() != previous_text:
                return True
        return False
