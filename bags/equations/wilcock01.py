"""Wilcock (2001) two-fraction calibrated bedload equation.

Reference
---------
Wilcock, P. R. (2001). Toward a practical method for estimating
    sediment-transport rates in gravel-bed rivers.
    *Earth Surface Processes and Landforms*, 26(13), 1395–1408.

Notes
-----
* Roughness: ``rough = 2 × D65`` (Manning-Strickler)
* Calibration: fits reference shear stresses TaurG and TaurS to
  observed gravel and sand bedload measurements.
* ``shields_stress`` is set to 0.0 in the result (no single normalising stress).
"""
from __future__ import annotations
import math
from typing import Sequence

from ..data import GrainSizeDistribution, ChannelGeometry, TransportResult
from .. import constants as C
from ..grain_size import percentile, gravel_sand_fractions
from ..hydraulics import (
    solve_depth_strickler, _xs_area_rh,
)


# ── transport sub-functions ───────────────────────────────────────────────────

def _hydraulics(Q: float, geom: ChannelGeometry, rough: float,
                tol: float = 1e-5) -> tuple[float, float, float]:
    """Return (H, ustar, area) using Manning-Strickler."""
    H, ustar, area = solve_depth_strickler(Q, geom, rough, tol=tol)
    if geom.cross_section is not None and geom.width is None:
        _, rh = _xs_area_rh(geom.cross_section, H)
        ustar = math.sqrt(C.g * rh * geom.slope)
    return H, ustar, area


def _gravel_transport(tau: float, Fg: float, TaurG: float,
                      ustar: float, area_or_width: float,
                      use_xs: bool) -> float:
    """Gravel volumetric transport rate (m³/s)."""
    if TaurG <= 0.0 or tau <= 0.0:
        return 0.0
    if use_xs:
        scale = Fg * area_or_width * C.g * ustar / (C.R * C.g)  # simplified
        # XS formula: Fg * Area * slope * ustar / R * W*(phi)
        # We factor out separately
        scale = Fg * area_or_width  # area
    else:
        scale = Fg * area_or_width  # width

    if tau > TaurG:
        w_star = 11.2 * (1.0 - 0.846 * TaurG / tau) ** 4.5
    else:
        w_star = 0.0025 * (tau / TaurG) ** 14.2

    if use_xs:
        # XS: scale * slope * ustar / R * w_star
        # BUT we need slope from the caller — pass it through ustar context
        return 0.0  # placeholder; see transport_rate for full impl
    return scale * ustar ** 3 / (C.R * C.g) * w_star


def _sand_transport(tau: float, Fs: float, TaurS: float,
                    ustar: float, area_or_width: float,
                    use_xs: bool) -> float:
    """Sand volumetric transport rate (m³/s)."""
    if TaurS <= 0.0 or tau <= 0.0:
        return 0.0
    dmy = 1.0 - 0.846 * math.sqrt(TaurS / tau)
    if dmy <= 0.0:
        return 0.0
    w_star = 11.2 * dmy ** 4.5
    if use_xs:
        return 0.0  # placeholder; see transport_rate for full impl
    return Fs * area_or_width * ustar ** 3 / (C.R * C.g) * w_star


# ── calibration ───────────────────────────────────────────────────────────────

def _sum_sq_gravel(TaurG: float, discharges: Sequence[float],
                   obs_gravel_kgs: Sequence[float],
                   geom: ChannelGeometry, rough: float, Fg: float,
                   tol: float) -> float:
    """Sum of squared log residuals for gravel transport."""
    S = geom.slope
    use_xs = geom.cross_section is not None and geom.width is None
    total = 0.0
    for Q, qg_obs in zip(discharges, obs_gravel_kgs):
        if qg_obs <= 0.0:
            continue
        H, ustar, area = _hydraulics(Q, geom, rough, tol=tol)
        if use_xs:
            _, rh = _xs_area_rh(geom.cross_section, H)
            ustar = math.sqrt(C.g * rh * S)
        tau = C.rho * ustar ** 2
        scale = Fg * (area if use_xs else geom.width)
        if tau > TaurG:
            w_star = 11.2 * (1.0 - 0.846 * TaurG / tau) ** 4.5
        else:
            w_star = 0.0025 * (tau / TaurG) ** 14.2
        if use_xs:
            qg_pred = scale * S * ustar / C.R * w_star * C.rho_s
        else:
            qg_pred = scale * ustar ** 3 / (C.R * C.g) * w_star * C.rho_s
        if qg_pred <= 0.0:
            continue
        total += (math.log(qg_obs) - math.log(qg_pred)) ** 2
    return total


def _sum_sq_sand(TaurS: float, discharges: Sequence[float],
                 obs_sand_kgs: Sequence[float],
                 geom: ChannelGeometry, rough: float, Fs: float,
                 tol: float) -> float:
    """Sum of squared log residuals for sand transport."""
    S = geom.slope
    use_xs = geom.cross_section is not None and geom.width is None
    total = 0.0
    for Q, qs_obs in zip(discharges, obs_sand_kgs):
        if qs_obs <= 0.0:
            continue
        H, ustar, area = _hydraulics(Q, geom, rough, tol=tol)
        if use_xs:
            _, rh = _xs_area_rh(geom.cross_section, H)
            ustar = math.sqrt(C.g * rh * S)
        tau = C.rho * ustar ** 2
        scale = Fs * (area if use_xs else geom.width)
        dmy = 1.0 - 0.846 * math.sqrt(TaurS / tau)
        if dmy <= 0.0:
            continue
        w_star = 11.2 * dmy ** 4.5
        if use_xs:
            qs_pred = scale * S * ustar / C.R * w_star * C.rho_s
        else:
            qs_pred = scale * ustar ** 3 / (C.R * C.g) * w_star * C.rho_s
        if qs_pred <= 0.0:
            continue
        total += (math.log(qs_obs) - math.log(qs_pred)) ** 2
    return total


def _golden_search_min(f, lo: float, hi: float, n_pts: int = 51,
                        tol: float = 1e-3) -> float:
    """Find minimum of f on a log-spaced grid [lo, hi], then refine."""
    while (hi - lo) / ((hi + lo) * 0.5 + 1e-30) > tol:
        pts = [lo * (hi / lo) ** (i / (n_pts - 1)) for i in range(n_pts)]
        vals = [f(x) for x in pts]
        idx = vals.index(min(vals))
        lo = pts[max(0, idx - 1)]
        hi = pts[min(n_pts - 1, idx + 1)]
    return (lo + hi) * 0.5


def calibrate(
    discharges: Sequence[float],
    observed_total_kgs: Sequence[float],
    observed_gravel_fraction: Sequence[float],
    geometry: ChannelGeometry,
    surface_gsd: GrainSizeDistribution,
    *,
    dk: float  = C.WILCOCK_DK,
    tol: float = 1e-5,
) -> tuple[float, float]:
    """Calibrate TaurG and TaurS to observed bedload samples.

    Parameters
    ----------
    discharges : Sequence[float]
        Measured water discharges (m³/s).
    observed_total_kgs : Sequence[float]
        Total measured bedload (kg/s) for each discharge.
    observed_gravel_fraction : Sequence[float]
        Gravel fraction (0–1) of total bedload for each sample.
    geometry : ChannelGeometry
    surface_gsd : GrainSizeDistribution
    dk : float
        Roughness multiplier (default 2.0).

    Returns
    -------
    (TaurG, TaurS) : tuple[float, float]
        Reference shear stresses for gravel (Pa) and sand (Pa).
    """
    D65_mm = percentile(surface_gsd, 65.0)
    rough  = dk * D65_mm / 1000.0
    Fg, Fs = gravel_sand_fractions(surface_gsd)

    obs_gravel = [qt * fg for qt, fg in zip(observed_total_kgs, observed_gravel_fraction)]
    obs_sand   = [qt * (1.0 - fg) for qt, fg in zip(observed_total_kgs, observed_gravel_fraction)]

    # Initial guesses
    TaurG0 = 0.04 * C.rho * C.R * C.g * rough
    TaurS0 = 0.1  * C.rho * C.R * C.g * 0.001  # ~1 mm grain

    TaurG = TaurG0
    TaurS = TaurS0

    if any(g > 0 for g in obs_gravel):
        f_g = lambda t: _sum_sq_gravel(t, discharges, obs_gravel, geometry, rough, Fg, tol)
        TaurG = _golden_search_min(f_g, TaurG0 / 10, TaurG0 * 10)

    if any(s > 0 for s in obs_sand):
        f_s = lambda t: _sum_sq_sand(t, discharges, obs_sand, geometry, rough, Fs, tol)
        TaurS = _golden_search_min(f_s, TaurS0 / 10, TaurS0 * 10)

    return TaurG, TaurS


# ── main transport function ───────────────────────────────────────────────────

def transport_rate(
    Q: float,
    geometry: ChannelGeometry,
    surface_gsd: GrainSizeDistribution,
    TaurG: float,
    TaurS: float,
    *,
    dk: float  = C.WILCOCK_DK,
    tol: float = 1e-5,
) -> TransportResult:
    """Compute gravel and sand bedload with the Wilcock (2001) two-fraction model.

    Parameters
    ----------
    Q : float
        Water discharge (m³/s).
    geometry : ChannelGeometry
    surface_gsd : GrainSizeDistribution
        Surface grain-size distribution (used for D65, Fg, Fs).
    TaurG : float
        Reference shear stress for gravel (Pa), from :func:`calibrate`.
    TaurS : float
        Reference shear stress for sand (Pa), from :func:`calibrate`.
    dk : float
        Roughness multiplier (default 2.0).

    Returns
    -------
    TransportResult
        ``bedload_by_fraction`` = [gravel_kgs, sand_kgs].
        ``shields_stress`` = 0.0 (no single normalising stress).
    """
    S = geometry.slope
    D65_mm = percentile(surface_gsd, 65.0)
    rough  = dk * D65_mm / 1000.0
    Fg, Fs = gravel_sand_fractions(surface_gsd)

    use_xs = geometry.cross_section is not None and geometry.width is None

    H, ustar, area = _hydraulics(Q, geometry, rough, tol=tol)
    if use_xs:
        _, rh = _xs_area_rh(geometry.cross_section, H)
        ustar = math.sqrt(C.g * rh * S)
    tau = C.rho * ustar ** 2

    width = geometry.width if not use_xs else None
    scale_g = Fg * (area if use_xs else width)
    scale_s = Fs * (area if use_xs else width)

    # gravel transport
    qg_vol = 0.0
    if TaurG > 0.0:
        if tau > TaurG:
            w_star_g = 11.2 * (1.0 - 0.846 * TaurG / tau) ** 4.5
        else:
            w_star_g = 0.0025 * (tau / TaurG) ** 14.2
        if use_xs:
            qg_vol = scale_g * S * ustar / C.R * w_star_g
        else:
            qg_vol = scale_g * ustar ** 3 / (C.R * C.g) * w_star_g

    # sand transport
    qs_vol = 0.0
    if TaurS > 0.0 and tau > 0.0:
        dmy = 1.0 - 0.846 * math.sqrt(TaurS / tau)
        if dmy > 0.0:
            w_star_s = 11.2 * dmy ** 4.5
            if use_xs:
                qs_vol = scale_s * S * ustar / C.R * w_star_s
            else:
                qs_vol = scale_s * ustar ** 3 / (C.R * C.g) * w_star_s

    qg_kgs = qg_vol * C.rho_s
    qs_kgs = qs_vol * C.rho_s
    total_kgs = qg_kgs + qs_kgs

    return TransportResult(
        discharge_m3s=Q,
        total_bedload_kgs=total_kgs,
        bedload_by_fraction=[qg_kgs, qs_kgs],
        shields_stress=0.0,
        flow_depth_m=H,
        shear_velocity_ms=ustar,
    )
