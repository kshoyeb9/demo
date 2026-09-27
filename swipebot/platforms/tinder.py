from swipebot.platforms.base import Platform


class Tinder(Platform):
    name = "tinder"
    url = "https://tinder.com/app/recs"
    card_selectors = [
        '[data-testid="swipe-card"]',
        '.recsCardboard__cards [data-keyboard-gamepad="true"][aria-hidden="false"]',
        ".recsCardboard__cards",
    ]
    stop_patterns = [
        r"you'?re out of likes",
        r"you'?ve run out of potential matches",
        r"there'?s no one new around you",
        r"log in with (google|facebook|phone)",
    ]
    popup_patterns = [
        r"it'?s a match",
        r"add tinder to your home screen",
        r"get tinder (gold|platinum|plus)",
        r"upgrade your like",
    ]

    def read_profile_text(self) -> str:
        # The card only shows a bio snippet; ArrowUp opens the full profile.
        self.page.keyboard.press("ArrowUp")
        self.page.wait_for_timeout(1_000)
        text = self.card_text()
        self.page.keyboard.press("ArrowDown")
        self.page.wait_for_timeout(500)
        return text
