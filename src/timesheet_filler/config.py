"""Configuration loading from YAML files."""

import yaml
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
    """Load configuration from a YAML file."""
    with open(path) as f:
        data = yaml.safe_load(f)

    labor_entries = [
        LaborEntry(
            org_index=str(entry["org_index"]),
            account_code=str(entry["account_code"]),
            percent=str(entry["percent"]),
        )
        for entry in data.get("labor", [])
    ]

    return PersonConfig(
        name=str(data["name"]),
        a_number=str(data["a_number"]),
        position_number=str(data["position_number"]),
        title=str(data["title"]),
        department=str(data["department"]),
        home_labor=str(data["home_labor"]),
        fte=str(data["fte"]),
        labor=labor_entries,
    )
