"""File with enums."""
import enum


class ReportGroupingType(enum.Enum):
    """Report grouping type."""
    FULL_MONTH = "full_month"    # For the entire month
    SPLIT_HALF = "split_half"    # 1–15 and 16–end