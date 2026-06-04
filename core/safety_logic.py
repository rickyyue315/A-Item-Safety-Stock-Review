"""Safety stock checking logic and constants."""
from openpyxl.styles import Font, PatternFill

SAFETY_ORDER = {
    "Aitem<25%": 0,
    "New item<15%": 1,
    "B/C/D item<15%": 2,
    "Aitem>75%": 3,
    "New item>100%": 4,
    "B/C/D item>100%": 5,
    "OK": 6,
    "Check": 7,
}

YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
AMBER_FILL = PatternFill(start_color="FFFFC000", end_color="FFFFC000", fill_type="solid")
DARK_RED_FONT = Font(color="8B0000")
PCT_FORMAT = "0.00%"
YELLOW_COLUMNS = {"Adj. Safety Stock", "Shop Monthly Sold Qty (P90D)", "Safety VS Monthly Sold Qty"}


def get_safety_checking(ratio, sasa_abc):
    """Determine safety checking category based on ratio and SASA ABC classification."""
    if sasa_abc == "A":
        if ratio is None or ratio < 0.25:
            return "Aitem<25%"
        if ratio > 0.75:
            return "Aitem>75%"
        return "OK"
    elif sasa_abc in ("B", "C", "D", "N", ""):
        if sasa_abc in ("N", ""):
            if ratio is None or ratio < 0.15:
                return "New item<15%"
            if ratio > 1.0:
                return "New item>100%"
            return "OK"
        else:
            if ratio is None or ratio < 0.15:
                return "B/C/D item<15%"
            if ratio > 1.0:
                return "B/C/D item>100%"
            return "OK"
    else:
        return "Check"