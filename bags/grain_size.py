"""Grain-size distribution statistics.

All functions work with a :class:`~bags.data.GrainSizeDistribution` whose
``sizes_mm`` values are in *ascending* order and ``finer_pct`` runs from
0 (or small) to 100 (or close).
"""
from __future__ import annotations
import math
from .data import GrainSizeDistribution


# ── helpers ───────────────────────────────────────────────────────────────────

def _psi(d_mm: float) -> float:
    """Convert grain diameter in mm to phi scale (log base-2)."""
    return math.log2(d_mm)


def _fractions(gsd: GrainSizeDistribution) -> tuple[list[float], list[float]]:
    """Return (psi_boundaries, f_i) aligned with size-class intervals.

    psi_boundaries has len(sizes_mm) entries.
    f_i[i] = |finer_pct[i+1] - finer_pct[i]| / 100  for i in 0..n-2
    """
    psi = [_psi(d) for d in gsd.sizes_mm]
    n = len(gsd.sizes_mm)
    f = [abs(gsd.finer_pct[i + 1] - gsd.finer_pct[i]) / 100.0
         for i in range(n - 1)]
    return psi, f


# ── public API ────────────────────────────────────────────────────────────────

def geometric_mean_phi(gsd: GrainSizeDistribution) -> tuple[float, float]:
    """Return (Dsg_m, std_phi): geometric-mean diameter (m) and arithmetic
    standard deviation in phi space.

    Matches VBA ``GetGeometricMeanGrainSizeAndArithmeticStandardDeviation``.
    """
    psi, f = _fractions(gsd)
    n = len(f)

    dsg_phi = sum(0.5 * (psi[i] + psi[i + 1]) * f[i] for i in range(n))
    var = sum((0.5 * (psi[i] + psi[i + 1]) - dsg_phi) ** 2 * f[i]
              for i in range(n))
    std_phi = math.sqrt(var)
    dsg_m = (2.0 ** dsg_phi) / 1000.0
    return dsg_m, std_phi


def percentile(gsd: GrainSizeDistribution, p: float) -> float:
    """Return the D_p grain size in mm via log-linear interpolation.

    *p* is the percentage finer (0–100).
    Matches VBA ``GetCharacteristicGrainSizeinMM``.
    """
    sizes = gsd.sizes_mm
    finer = gsd.finer_pct
    n = len(sizes)

    # clamp at edges
    if p <= finer[0]:
        return sizes[0]
    if p >= finer[-1]:
        return sizes[-1]

    for i in range(n - 1):
        lo, hi = finer[i], finer[i + 1]
        if lo <= p <= hi:
            if hi == lo:
                return sizes[i]
            t = (p - lo) / (hi - lo)
            # log-linear interpolation
            return sizes[i] * (sizes[i + 1] / sizes[i]) ** t
    return sizes[-1]


def gravel_sand_fractions(gsd: GrainSizeDistribution) -> tuple[float, float]:
    """Return (Fg, Fs): fraction coarser/finer than 2 mm.

    Fg + Fs = 1.0.
    Interpolates the % finer at 2 mm from the GSD.
    """
    pct_finer_at_2mm = percentile.__wrapped__(gsd, 2.0) if hasattr(
        percentile, '__wrapped__') else _pct_at_size(gsd, 2.0)
    Fs = pct_finer_at_2mm / 100.0
    Fg = 1.0 - Fs
    return Fg, Fs


def _pct_at_size(gsd: GrainSizeDistribution, target_mm: float) -> float:
    """Interpolate % finer at *target_mm* from the GSD."""
    sizes = gsd.sizes_mm
    finer = gsd.finer_pct
    n = len(sizes)

    if target_mm <= sizes[0]:
        return finer[0]
    if target_mm >= sizes[-1]:
        return finer[-1]

    for i in range(n - 1):
        if sizes[i] <= target_mm <= sizes[i + 1]:
            if sizes[i + 1] == sizes[i]:
                return finer[i]
            t = math.log(target_mm / sizes[i]) / math.log(sizes[i + 1] / sizes[i])
            return finer[i] + t * (finer[i + 1] - finer[i])
    return finer[-1]


# fix the reference above
def gravel_sand_fractions(gsd: GrainSizeDistribution) -> tuple[float, float]:  # noqa: F811
    """Return (Fg, Fs): fraction coarser/finer than 2 mm."""
    pct_finer_at_2mm = _pct_at_size(gsd, 2.0)
    Fs = pct_finer_at_2mm / 100.0
    Fg = 1.0 - Fs
    return Fg, Fs


def fraction_volumes(gsd: GrainSizeDistribution) -> tuple[list[float], list[float], list[float]]:
    """Return (psi_boundaries, Di_m, f_i) for use in transport equations.

    *psi_boundaries* — len(sizes_mm) phi values.
    *Di_m*           — geometric-mean diameter per size class in metres.
    *f_i*            — fraction by weight in each class.
    """
    psi, f = _fractions(gsd)
    n = len(f)
    di_m = [(2.0 ** (0.5 * (psi[i] + psi[i + 1]))) / 1000.0 for i in range(n)]
    return psi, di_m, f
