"""Canonical WFMHub report-design tokens.

Business calculations must never depend on this module.  It is the single code
authority for workbook colors and the versioned presentation contract.
"""

from __future__ import annotations


REPORT_DESIGN_ID = "WFMHUB-DESIGN"
REPORT_DESIGN_VERSION = "2.2.0"

TITLE_FONT = "Aptos Display"
BODY_FONT = "Aptos"

COLORS = {
    "dark": "#0B1F33",
    "navy": "#0B1F33",
    "muted": "#536474",
    "teal": "#007C83",
    "teal_light": "#DFF3F3",
    "canvas": "#F4F7F9",
    "ink": "#1F2933",
    "gold": "#D6A84B",
    "rule": "#9AA6B2",
    "thin": "#D8E0E6",
    "line": "#D8E0E6",
    "blue": "#0563C1",
    "blue_light": "#E3F0FA",
    "green": "#1F7A53",
    "green_light": "#DDF3E8",
    "amber": "#A65F00",
    "amber_light": "#FFF1CC",
    "red": "#B42318",
    "red_light": "#FDE7E5",
    "purple": "#6E56CF",
    "purple_light": "#EEEAFE",
    "future": "#9AA6B2",
    "future_light": "#EEF1F4",
    "white": "#FFFFFF",
}
