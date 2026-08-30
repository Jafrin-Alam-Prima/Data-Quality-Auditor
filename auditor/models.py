"""Data structures shared by the audit engine and the feedback generator."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Issue:
    """One finding: a single problem type in a single column or sheet."""

    section: str                 # heading it appears under in the feedback
    code: str                    # machine name: "blank", "duplicate", ...
    title: str                   # short human label for the results screen
    sheet: str = ""              # sheet the problem is in
    column: str = ""             # column the problem is in
    count: int = 0               # how many cells / rows are affected
    values: list[str] = field(default_factory=list)   # distinct offending values
    rows: list[int] = field(default_factory=list)     # Excel row numbers
    note: str = ""               # extra context for the results screen only

    @property
    def examples(self) -> list[str]:
        return self.values[:5]


@dataclass
class ReferenceList:
    """A validated list of allowed values taken from a reference sheet."""

    key: str
    label: str                                   # "Supplier sheet"
    sheet_name: str = ""
    values: list[str] = field(default_factory=list)          # canonical values
    exact: set[str] = field(default_factory=set)             # for exact matching
    normalized: dict[str, str] = field(default_factory=dict) # lower-case -> canonical
    # Short forms of a canonical value that carries a parenthetical
    # explanation, e.g. "good" -> "Good (No visible damage, no repairs
    # completed)". Kept separate from `normalized` because a match here is a
    # genuinely different (shorter) string, not a formatting variant of the
    # same text — it is accepted silently, never reported as "spelled
    # differently".
    aliases: dict[str, str] = field(default_factory=dict)
    available: bool = True                       # False when the sheet is missing

    def match(self, key: str) -> str | None:
        """Return the canonical value for a normalised key, or None."""
        return self.normalized.get(key) or self.aliases.get(key)

    def contains(self, key: str) -> bool:
        """True when `key` matches this reference exactly or as a short form."""
        return key in self.normalized or key in self.aliases


@dataclass
class AuditResult:
    filename: str
    country: str
    sheet_overview: list[dict] = field(default_factory=list)
    resolved_sheets: dict[str, str] = field(default_factory=dict)
    missing_sheets: list[str] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    asset_row_count: int = 0

    def by_section(self) -> dict[str, list[Issue]]:
        from .config import SECTION_ORDER

        grouped: dict[str, list[Issue]] = {}
        for issue in self.issues:
            grouped.setdefault(issue.section, []).append(issue)
        order = {name: index for index, name in enumerate(SECTION_ORDER)}
        return dict(sorted(grouped.items(), key=lambda kv: order.get(kv[0], 999)))

    @property
    def has_issues(self) -> bool:
        return bool(self.issues)
