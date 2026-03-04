from __future__ import annotations
from dataclasses import dataclass, field


@dataclass
class GrainSizeDistribution:
    """Cumulative grain-size distribution."""
    sizes_mm: list[float]   # grain diameters in mm, ascending
    finer_pct: list[float]  # cumulative % finer (0–100)


@dataclass
class ChannelGeometry:
    """Channel geometry and hydraulic parameters."""
    slope: float                        # dimensionless (m/m)
    width: float | None = None          # bankfull width in m (None → use cross_section)
    cross_section: list[tuple[float, float]] | None = None  # (station_m, elevation_m) pairs
    mannings_n: float | None = None          # main-channel Manning's n
    mannings_n_left: float | None = None     # left-floodplain Manning's n
    mannings_n_right: float | None = None    # right-floodplain Manning's n


@dataclass
class TransportResult:
    """Output from a single-discharge bedload calculation."""
    discharge_m3s: float
    total_bedload_kgs: float          # kg/s
    bedload_by_fraction: list[float]  # kg/s per size class (aligned with GSD fractions)
    shields_stress: float             # normalised Shields stress (equation-specific)
    flow_depth_m: float               # max water depth (m)
    shear_velocity_ms: float          # bed shear velocity (m/s)
