from swipebot.platforms.base import Platform
from swipebot.platforms.bumble import Bumble
from swipebot.platforms.tinder import Tinder

PLATFORMS: dict[str, type[Platform]] = {"tinder": Tinder, "bumble": Bumble}
