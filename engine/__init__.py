"""CCTS engine modules.

Modular refactor of the CCC Monte Carlo + Market Price Discovery app:
- data_pipeline.py   : Excel/CSV ingestion, Main Sheet rebuild, sector mapping, data isolation
- ccc_sim.py         : 10,000-iteration CCC Monte Carlo
- price_discovery.py : 5,000-session market clearing (pluggable engines)
- macc_cbam.py       : MACC parsing, plant-scaled abatement, EU/UK CBAM calculator
"""

from .data_pipeline import (
    load_main_sheet_from_excel,
    rebuild_main_sheet_from_sources,
    apply_sidebar_adjustments,
    clean_num,
    load_macc,
    scale_macc_to_plant,
    get_sector_emissions_mtco2,
    add_sector_column,
    isolate_for_company,
    mask_competitors,
)
from .ccc_sim import run_ccc_monte_carlo
from .price_discovery import (
    run_price_discovery,
    ENGINE_RANDOM,
    ENGINE_MACC,
)
from .macc_cbam import (
    compute_plant_macc_table,
    abate_vs_buy_decision,
    compute_cbam_liability,
)

__all__ = [
    "load_main_sheet_from_excel",
    "rebuild_main_sheet_from_sources",
    "apply_sidebar_adjustments",
    "clean_num",
    "load_macc",
    "scale_macc_to_plant",
    "get_sector_emissions_mtco2",
    "add_sector_column",
    "isolate_for_company",
    "mask_competitors",
    "run_ccc_monte_carlo",
    "run_price_discovery",
    "ENGINE_RANDOM",
    "ENGINE_MACC",
    "compute_plant_macc_table",
    "abate_vs_buy_decision",
    "compute_cbam_liability",
]
