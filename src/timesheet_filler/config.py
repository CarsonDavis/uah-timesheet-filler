"""Configuration loading from TOML files."""

import tomllib
from dataclasses import dataclass
from pathlib import Path


@dataclass
class LaborEntry:
    """A single labor distribution entry."""

    org_index: str
    account_code: str
    percent: str


@dataclass
class PersonConfig:
    """Person configuration from details.toml."""

    name: str
    a_number: str
    position_number: str
    title: str
    department: str
    home_labor: str
    fte: str
    labor: list[LaborEntry]

    @property
    def last_name(self) -> str:
        """Extract last name (last word) from full name."""
        return self.name.split()[-1]


def load_config(path: Path) -> PersonConfig:
    """Load configuration from a TOML file."""
    with open(path, "rb") as f:
        data = tomllib.load(f)

    person = data["person"]
    labor_entries = [
        LaborEntry(
            org_index=entry["org_index"],
            account_code=entry["account_code"],
            percent=entry["percent"],
        )
        for entry in data.get("labor", [])
    ]

    return PersonConfig(
        name=person["name"],
        a_number=person["a_number"],
        position_number=person["position_number"],
        title=person["title"],
        department=person["department"],
        home_labor=person["home_labor"],
        fte=person["fte"],
        labor=labor_entries,
    )
