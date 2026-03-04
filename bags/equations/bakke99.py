"""Bakke et al. (1999) calibrated Parker-Klingeman bedload equation.

Reference
---------
Bakke, P. D., Basdekas, P. O., Booth, D. B., & Sibley, P. K. (1999).
    Calibrated sediment transport model for a low-gradient gravel-bed stream.
    *Environmental Management*, 24(4), 463-478.

Notes
-----
The calibration procedure (originally described in Bakke et al. 1999)
alternately optimises:
1. Reference Shields stress ``taur50`` with fixed ``beta``.
2. Hiding exponent ``beta`` with fixed ``taur50``.
Cycling is repeated up to ``max_iterations`` times.

After calibration, predictions use the Parker-Klingeman (1982) multi-fraction
transport function (``use_pkm`` kwarg selects the PKM variant).
"""
from __future__ import annotations
import math
from typing import Sequence

from ..data import GrainSizeDistribution, ChannelGeometry, TransportResult
from .. import constants as C
from ..grain_size import percentile, fraction_volumes
from ..hydraulics import solve_depth_loglaw, _xs_area_rh
from .parker82 import pk_g, pkm_g


# ── internal helpers ──────────────────────────────────────────────────────────

def _pk_transport_vol(Q: float, geom: ChannelGeometry, D50_m: float,
                      rough: float, taur50: float, beta: float,
                      psi: list, Di_m: list, f_frac: list,
                      use_pkm: bool, tol: float) -> tuple[float, float, float, list]:
    """Return (H, ustar, qs_vol, p_norm) for one discharge."""
    S = geom.slope
    g_func = pkm_g if use_pkm else pk_g

    H, ustar, area = solve_depth_loglaw(Q, geom, rough, tol=tol)
    if geom.cross_section is not None and geom.width is None:
        _, rh_val = _xs_area_rh(geom.cross_section, H)
        ustar_phi = math.sqrt(C.g * rh_val * S)
    else:
        ustar_phi = ustar

    phi50  = ustar_phi ** 2 / (C.R * C.g * D50_m * taur50)
    n_frac = len(f_frac)

    qs_sum = 0.0
    p = []
    for i in range(n_frac):
        dmy = phi50 / (Di_m[i] / D50_m) ** beta
        gi  = g_func(dmy) * f_frac[i]
        p.append(gi)
        qs_sum += gi

    p_norm = [pi / qs_sum for pi in p] if qs_sum > 0.0 else [0.0] * n_frac

    if geom.width is not None and geom.cross_section is None:
        qs_vol = qs_sum * geom.width * ustar ** 3 / (C.R * C.g)
    else:
        qs_vol = qs_sum * ustar * S * area / C.R

    return H, ustar, qs_vol, p_norm


def _sum_sq_total(taur50: float, beta: float,
                  discharges: Sequence[float], obs_kgs: Sequence[float],
                  geom: ChannelGeometry, D50_m: float, rough: float,
                  psi, Di_m, f_frac, use_pkm: bool, tol: float) -> float:
    """Sum of squared log residuals for total bedload."""
    total = 0.0
    for Q, qo in zip(discharges, obs_kgs):
        if qo <= 0.0:
            continue
        _, _, qs_vol, _ = _pk_transport_vol(Q, geom, D50_m, rough,
                                             taur50, beta, psi, Di_m,
                                             f_frac, use_pkm, tol)
        qs_kgs = qs_vol * C.rho_s
        if qs_kgs <= 0.0:
            continue
        total += (math.log(qo) - math.log(qs_kgs)) ** 2
    return total


def _bedload_d50(p_norm: list[float], Di_m: list[float]) -> float:
    """Return bedload D50 (m) from normalised size-class fractions."""
    # cumulative % finer
    cum = 0.0
    for i, pi in enumerate(p_norm):
        prev = cum
        cum += pi
        if cum >= 0.5:
            if cum == prev:
                return Di_m[i]
            frac = (0.5 - prev) / (cum - prev)
            if i == 0:
                return Di_m[0] * (Di_m[0] / Di_m[0]) ** frac  # edge
            return Di_m[i - 1] * (Di_m[i] / Di_m[i - 1]) ** frac
    return Di_m[-1]


def _sum_sq_d50(beta: float, taur50: float,
                discharges: Sequence[float], obs_d50_m: Sequence[float],
                geom: ChannelGeometry, D50_m: float, rough: float,
                psi, Di_m, f_frac, use_pkm: bool, tol: float) -> float:
    """Sum of squared log residuals for bedload D50."""
    total = 0.0
    for Q, d50_obs in zip(discharges, obs_d50_m):
        if d50_obs is None or d50_obs <= 0.0:
            continue
        _, _, qs_vol, p_norm = _pk_transport_vol(Q, geom, D50_m, rough,
                                                  taur50, beta, psi, Di_m,
                                                  f_frac, use_pkm, tol)
        if qs_vol <= 0.0:
            continue
        d50_pred = _bedload_d50(p_norm, Di_m)
        if d50_pred <= 0.0:
            continue
        total += (math.log(d50_obs) - math.log(d50_pred)) ** 2
    return total


def _golden_min(f, lo: float, hi: float, n_pts: int = 51,
                tol: float = 1e-3) -> float:
    """Find minimum of f on log-spaced grid, refine until convergence."""
    while (hi - lo) / ((hi + lo) * 0.5 + 1e-30) > tol:
        pts = [lo * (hi / lo) ** (i / (n_pts - 1)) for i in range(n_pts)]
        vals = [f(x) for x in pts]
        idx  = vals.index(min(vals))
        lo   = pts[max(0, idx - 1)]
        hi   = pts[min(n_pts - 1, idx + 1)]
    return (lo + hi) * 0.5


# ── public API ────────────────────────────────────────────────────────────────

def calibrate(
    discharges: Sequence[float],
    observed_kgs: Sequence[float],
    geometry: ChannelGeometry,
    substrate_gsd: GrainSizeDistribution,
    *,
    observed_d50_mm: Sequence[float] | None = None,
    use_pkm: bool    = False,
    dk: float        = C.PK82_DK,
    max_iterations: int = 6,
    tol: float       = 1e-5,
) -> tuple[float, float]:
    """Calibrate (taur50, beta) to observed bedload measurements.

    Parameters
    ----------
    discharges : Sequence[float]
        Measured water discharges (m³/s).
    observed_kgs : Sequence[float]
        Total measured bedload (kg/s) for each discharge.
    geometry : ChannelGeometry
    substrate_gsd : GrainSizeDistribution
        Substrate grain-size distribution.
    observed_d50_mm : Sequence[float] | None
        Observed bedload D50 (mm) for each discharge sample.  When provided,
        ``beta`` is calibrated by minimising log-residuals of predicted vs
        observed bedload D50 (matching the VBA ``GoGetExponent`` step).
        When ``None``, ``beta`` is calibrated against total bedload instead.
    use_pkm : bool
        Use PKM G function (default: False → PK G function).
    dk : float
        Roughness multiplier (default 10.7).
    max_iterations : int
        Maximum alternating optimisation cycles (default 6).

    Returns
    -------
    (taur50, beta) : tuple[float, float]
    """
    D50_mm = percentile(substrate_gsd, 50.0)
    D50_m  = D50_mm / 1000.0
    rough  = dk * D50_m
    psi, Di_m, f_frac = fraction_volumes(substrate_gsd)

    taur50  = C.PK82_TAUR50
    beta    = C.PK82_BETA

    obs_d50_m: Sequence[float] | None = None
    if observed_d50_mm is not None:
        obs_d50_m = [d / 1000.0 for d in observed_d50_mm]

    for iteration in range(max_iterations):
        # optimise taur50 with fixed beta
        f_t = lambda t: _sum_sq_total(t, beta, discharges, observed_kgs,
                                       geometry, D50_m, rough, psi, Di_m,
                                       f_frac, use_pkm, tol)
        lo = taur50 / (5.0 if iteration == 0 else 2.0)
        hi = taur50 * (5.0 if iteration == 0 else 2.0)
        taur50_new = _golden_min(f_t, lo, hi)

        # optimise beta with new taur50
        lo = beta / (5.0 if iteration == 0 else 3.0)
        hi = beta * (5.0 if iteration == 0 else 3.0)
        if obs_d50_m is not None:
            f_b = lambda b: _sum_sq_d50(b, taur50_new, discharges, obs_d50_m,
                                         geometry, D50_m, rough, psi, Di_m,
                                         f_frac, use_pkm, tol)
        else:
            f_b = lambda b: _sum_sq_total(taur50_new, b, discharges, observed_kgs,
                                           geometry, D50_m, rough, psi, Di_m,
                                           f_frac, use_pkm, tol)
        beta_new = _golden_min(f_b, lo, hi)

        # convergence check
        dt = abs(taur50_new - taur50) / taur50 if taur50 > 0 else 1.0
        db = abs(beta_new  - beta)   / beta   if beta   > 0 else 1.0
        taur50 = taur50_new
        beta   = beta_new

        if dt < 0.001 and db < 0.001 and iteration >= 2:
            break

    return taur50, beta


def transport_rate(
    Q: float,
    geometry: ChannelGeometry,
    substrate_gsd: GrainSizeDistribution,
    taur50: float,
    beta: float,
    *,
    use_pkm: bool = False,
    dk: float     = C.PK82_DK,
    tol: float    = 1e-5,
) -> TransportResult:
    """Compute bedload transport with Bakke et al. (1999) calibrated parameters.

    Parameters
    ----------
    Q : float
        Water discharge (m³/s).
    geometry : ChannelGeometry
    substrate_gsd : GrainSizeDistribution
    taur50 : float
        Calibrated reference Shields stress (from :func:`calibrate`).
    beta : float
        Calibrated hiding-function exponent (from :func:`calibrate`).
    use_pkm : bool
        Use PKM G function (default: False → PK).
    dk : float
        Roughness multiplier (default 10.7).

    Returns
    -------
    TransportResult
    """
    D50_mm = percentile(substrate_gsd, 50.0)
    D50_m  = D50_mm / 1000.0
    rough  = dk * D50_m
    psi, Di_m, f_frac = fraction_volumes(substrate_gsd)

    H, ustar, qs_vol, p_norm = _pk_transport_vol(
        Q, geometry, D50_m, rough, taur50, beta,
        psi, Di_m, f_frac, use_pkm, tol)

    # phi50 for reporting
    if geometry.cross_section is not None and geometry.width is None:
        _, rh_val = _xs_area_rh(geometry.cross_section, H)
        ustar_phi = math.sqrt(C.g * rh_val * geometry.slope)
    else:
        ustar_phi = ustar
    phi50 = ustar_phi ** 2 / (C.R * C.g * D50_m * taur50)

    qs_kgs = qs_vol * C.rho_s
    fractions_kgs = [qs_kgs * p for p in p_norm]

    return TransportResult(
        discharge_m3s=Q,
        total_bedload_kgs=qs_kgs,
        bedload_by_fraction=fractions_kgs,
        shields_stress=phi50,
        flow_depth_m=H,
        shear_velocity_ms=ustar,
    )
