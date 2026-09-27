"""Turns the visible text of a profile card into structured fields.

Both apps render name/age/distance as plain text, and their CSS class names
are obfuscated and change often, so parsing the card's text is more durable
than depending on specific elements.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

MILES_TO_KM = 1.609344

_NAME_AGE = re.compile(r"^\s*([^\d\n,]{1,40}?)\s*,?\s+(\d{2})\s*$", re.MULTILINE)
_DISTANCE = re.compile(
    r"(less than\s+)?(\d+(?:\.\d+)?)\s*(km|kilometers?|mi|miles?)\s+away", re.IGNORECASE
)

# UI chrome that shows up in the card text but isn't part of the bio.
_NOISE = re.compile(
    r"^(like|nope|pass|super ?like|rewind|boost|open profile|report|block|share"
    r"|recently active|active today|lives in .*|\d+(\.\d+)? ?(km|mi|miles?) away"
    r"|less than .* away|verified|show more|show less|about me|my basics|looking for)$",
    re.IGNORECASE,
)


@dataclass
class Profile:
    name: str | None
    age: int | None
    distance_km: float | None
    bio: str
    raw_text: str

    @property
    def readable(self) -> bool:
        return bool(self.name or self.age or self.bio)


def parse_profile(text: str) -> Profile:
    text = (text or "").strip()
    name = age = None
    name_line = None
    if m := _NAME_AGE.search(text):
        name, age = m.group(1).strip(), int(m.group(2))
        name_line = m.group(0).strip()

    distance_km = None
    if m := _DISTANCE.search(text):
        value = float(m.group(2))
        distance_km = value * MILES_TO_KM if m.group(3).lower().startswith("mi") else value

    bio_lines = []
    for line in text.splitlines():
        line = line.strip()
        if not line or line == name_line or _NOISE.match(line):
            continue
        bio_lines.append(line)

    return Profile(name=name, age=age, distance_km=distance_km, bio="\n".join(bio_lines), raw_text=text)
