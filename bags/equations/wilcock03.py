"""Wilcock & Crowe (2003) surface-based mixed-size bedload equation.

Reference
---------
Wilcock, P. R., & Crowe, J. C. (2003). Surface-based transport model for
    mixed-size sediment.
    *Journal of Hydraulic Engineering*, 129(2), 120–128.

Notes
-----
* Roughness: ``rough = 2 × D65`` (Manning-Strickler)
* Hydraulics: Manning-Strickler explicit formula (constant width) or bisection
  (cross-section).
* ``shields_stress`` in the result is τ / τ*_rsg (normalised to the geometric
  mean reference stress).
"""
from __future__ import annotations
import math

from ..data import GrainSizeDistribution, ChannelGeometry, TransportResult
from .. import constants as C
from ..grain_size import geometric_mean_phi, percentile, gravel_sand_fractions, fraction_volumes
from ..hydraulics import (
    solve_depth_strickler, _xs_area_rh, mannings_n_correction,
)


# ── sub-functions ─────────────────────────────────────────────────────────────

def reference_stress(Dsg_m: float, Fs: float) -> float:
    """Reference Shields stress τ*_rsg (Pa).

    Parameters
    ----------
    Dsg_m : float  — geometric mean grain size (m)
    Fs    : float  — sand fraction (finer than 2 mm)
    """
    return C.rho * C.R * C.g * Dsg_m * (0.021 + 0.015 / math.exp(20.0 * Fs))


def hiding_coeff(Di_m: float, Dsg_m: float) -> float:
    """Hiding/exposure exponent *b_i*."""
    return 0.67 / (1.0 + math.exp(1.5 - Di_m / Dsg_m))


def reference_stress_i(Di_m: float, Dsg_m: float, taursg: float) -> float:
    """Reference Shields stress for size class *i* (Pa)."""
    b = hiding_coeff(Di_m, Dsg_m)
    return (Di_m / Dsg_m) ** b * taursg


def wilcock_g(phi: float) -> float:
    """Wilcock & Crowe (2003) transport function W*."""
    if phi > 1.35:
        return 14.0 * (1.0 - 0.894 / math.sqrt(phi)) ** 4.5
    return 0.002 * phi ** 7.5


# ── main transport function ───────────────────────────────────────────────────

def transport_rate(
    Q: float,
    geometry: ChannelGeometry,
    surface_gsd: GrainSizeDistribution,
    *,
    dk: float   = C.WILCOCK_DK,
    tol: float  = 1e-5,
) -> TransportResult:
    """Compute bedload transport with Wilcock & Crowe (2003).

    Parameters
    ----------
    Q : float
        Water discharge (m³/s).
    geometry : ChannelGeometry
        Channel geometry.
    surface_gsd : GrainSizeDistribution
        Surface grain-size distribution.
    dk : float
        Roughness multiplier: ``rough = dk × D65`` (default 2.0).

    Returns
    -------
    TransportResult
        Bedload transport result; ``total_bedload_kgs`` in kg/s,
        ``shields_stress`` = τ / τ*_rsg.
    """
    S = geometry.slope

    Dsg_m, _ = geometric_mean_phi(surface_gsd)
    D65_mm    = percentile(surface_gsd, 65.0)
    D65_m     = D65_mm / 1000.0
    rough     = dk * D65_m

    Fg, Fs = gravel_sand_fractions(surface_gsd)
    taursg  = reference_stress(Dsg_m, Fs)

    psi, Di_m, f = fraction_volumes(surface_gsd)
    n_frac = len(f)

    # ── hydraulics ────────────────────────────────────────────────────────────
    corr = None
    if geometry.mannings_n is not None:
        nD = 0.04 * rough ** (1.0 / 6.0)
        if geometry.mannings_n > nD:
            corr = mannings_n_correction(Q, geometry, rough,
                                         geometry.mannings_n,
                                         solver="strickler", tol=tol)

    if corr is not None:
        H, ustar, area = corr
        tau = C.rho * ustar ** 2
    else:
        H, ustar, area = solve_depth_strickler(Q, geometry, rough, tol=tol)
        if geometry.cross_section is not None and geometry.width is None:
            _, rh_val = _xs_area_rh(geometry.cross_section, H)
            ustar = math.sqrt(C.g * rh_val * S)
        tau  = C.rho * ustar ** 2

    # ── transport calculation ─────────────────────────────────────────────────
    phi_sg = tau / taursg   # normalised transport stage

    qs_sum = 0.0
    p = []
    for i in range(n_frac):
        tauri  = reference_stress_i(Di_m[i], Dsg_m, taursg)
        phi_i  = tau / tauri
        gi     = wilcock_g(phi_i) * f[i]
        p.append(gi)
        qs_sum += gi

    if qs_sum > 0.0:
        p_norm = [pi / qs_sum for pi in p]
    else:
        p_norm = [0.0] * n_frac

    if geometry.width is not None and geometry.cross_section is None:
        qs_vol = qs_sum * geometry.width * ustar ** 3 / (C.R * C.g)
    else:
        qs_vol = qs_sum * area * S * ustar / C.R

    qs_kgs = qs_vol * C.rho_s
    fractions_kgs = [qs_kgs * p_norm[i] for i in range(n_frac)]

    return TransportResult(
        discharge_m3s=Q,
        total_bedload_kgs=qs_kgs,
        bedload_by_fraction=fractions_kgs,
        shields_stress=phi_sg,
        flow_depth_m=H,
        shear_velocity_ms=ustar,
    )
