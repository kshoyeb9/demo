from swipebot.platforms.base import Platform


class Bumble(Platform):
    name = "bumble"
    url = "https://bumble.com/app"
    card_selectors = [
        '[data-testid="encounters-story"]',
        ".encounters-story-profile",
        ".encounters-album",
    ]
    stop_patterns = [
        r"you'?ve run out of (swipes|votes)",
        r"that'?s everyone( for now)?",
        r"no one new( nearby)?",
        r"continue with (google|facebook|apple)",
    ]
    popup_patterns = [
        r"it'?s a match|you matched",
        r"try bumble (boost|premium)",
        r"extend your search",
    ]
