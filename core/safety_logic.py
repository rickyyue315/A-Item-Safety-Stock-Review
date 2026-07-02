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


# Thresholds for the "A item overstocked but D001 nearly empty" check.
# Vendor items do not flow through D001, so this only applies to Warehouse items.
AITEM_OVER_RATIO = 2.0       # Safety/Monthly ratio > 200%
D001_LOW_STOCK = 20          # D001 Net Stock Qty below this many pcs

AITEM_OVER_RATIO_TIER2 = 1.5   # Safety/Monthly ratio > 150%
D001_LOW_STOCK_TIER2 = 50      # D001 Net Stock Qty below this many pcs


def is_aitem_overstock_short_d001(ratio, item_class_val, supply_source, d001_net_stock):
    """Detect A-item SKUs that are overstocked (ratio > 200%) while D001 is nearly
    empty (< 20 pcs).

    Only meaningful for Warehouse supply source — Vendor items do not go through
    D001. Returns True when all conditions hold, False otherwise.
    """
    sasa_abc = str(item_class_val)[0] if item_class_val else ""
    return (
        sasa_abc == "A"
        and ratio is not None
        and ratio > AITEM_OVER_RATIO
        and str(supply_source or "").strip().lower() == "warehouse"
        and d001_net_stock is not None
        and d001_net_stock < D001_LOW_STOCK
    )


def is_aitem_overstock_short_d001_tier2(ratio, item_class_val, supply_source, d001_net_stock):
    """Detect A-item SKUs that are overstocked (ratio > 150%) while D001 is low
    (< 50 pcs). Broader tier that includes the stricter >200% & <20pcs SKUs.

    Only meaningful for Warehouse supply source — Vendor items do not go through
    D001. Returns True when all conditions hold, False otherwise.
    """
    sasa_abc = str(item_class_val)[0] if item_class_val else ""
    return (
        sasa_abc == "A"
        and ratio is not None
        and ratio > AITEM_OVER_RATIO_TIER2
        and str(supply_source or "").strip().lower() == "warehouse"
        and d001_net_stock is not None
        and d001_net_stock < D001_LOW_STOCK_TIER2
    )