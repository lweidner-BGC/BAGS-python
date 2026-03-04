"""Parker-Klingeman (1982) and Parker-Klingeman-McLean (1982) substrate-based equations.

References
----------
Parker, G., & Klingeman, P. C. (1982). On why gravel bed streams are paved.
    *Water Resources Research*, 18(5), 1409–1423.

Parker, G., Klingeman, P. C., & McLean, D. G. (1982). Bedload and size
    distribution in paved gravel-bed streams.
    *Journal of Hydraulic Engineering*, 108(4), 544–571.

Notes
-----
Two variants are implemented:

``use_pkm=False``  — Parker-Klingeman (1982) PK variant.
    G function: if φ > 0.95 → 11.2·(1 - 0.853/φ)^4.5, else 0.00242947·φ^35.714

``use_pkm=True``  — Parker-Klingeman-McLean (1982) PKM variant.
    G function (W*): if φ ≤ 0.95 → 0.006518·φ^32.978,
                      if φ ≤ 1.65 → 0.0025·exp(14.2·(φ-1)-9.28·(φ-1)²),
                      else         → 11.2·(1-0.822/φ)^4.5
"""
from __future__ import annotations
import math

from ..data import GrainSizeDistribution, ChannelGeometry, TransportResult
from .. import constants as C
from ..grain_size import percentile, fraction_volumes
from ..hydraulics import (
    solve_depth_loglaw, _xs_area_rh, mannings_n_correction,
)


# ── G / W* functions ─────────────────────────────────────────────────────────

def pk_g(phi: float) -> float:
    """Parker-Klingeman (1982) PK transport function."""
    if phi > 0.95:
        return 11.2 * (1.0 - 0.853 / phi) ** 4.5
    return 0.00242947 * phi ** 35.71387887


def pkm_g(phi: float) -> float:
    """Parker-Klingeman-McLean (1982) W* transport function."""
    if phi <= 0.95:
        return 0.006518 * phi ** 32.978
    if phi <= 1.65:
        return 0.0025 * math.exp(14.2 * (phi - 1.0) - 9.28 * (phi - 1.0) ** 2)
    return 11.2 * (1.0 - 0.822 / phi) ** 4.5


# ── main transport function ───────────────────────────────────────────────────

def transport_rate(
    Q: float,
    geometry: ChannelGeometry,
    substrate_gsd: GrainSizeDistribution,
    *,
    use_pkm: bool  = False,
    taur50: float  = C.PK82_TAUR50,
    beta: float    = C.PK82_BETA,
    dk: float      = C.PK82_DK,
    tol: float     = 1e-5,
) -> TransportResult:
    """Compute bedload transport with Parker-Klingeman (1982).

    Parameters
    ----------
    Q : float
        Water discharge (m³/s).
    geometry : ChannelGeometry
        Channel geometry.
    substrate_gsd : GrainSizeDistribution
        Substrate (sub-surface) grain-size distribution.
    use_pkm : bool
        If True, use the PKM variant (different G function).
    taur50 : float
        Reference Shields stress based on D50 (default 0.0876).
    beta : float
        Hiding-function exponent (default 0.018).
    dk : float
        Roughness multiplier: ``rough = dk × D50`` (default 10.7).

    Returns
    -------
    TransportResult
        Bedload transport result; ``total_bedload_kgs`` in kg/s.
    """
    S = geometry.slope
    g_func = pkm_g if use_pkm else pk_g

    D50_mm = percentile(substrate_gsd, 50.0)
    D50_m  = D50_mm / 1000.0
    rough  = dk * D50_m

    psi, Di_m, f = fraction_volumes(substrate_gsd)
    n_frac = len(f)

    # ── hydraulics ────────────────────────────────────────────────────────────
    corr = None
    if geometry.mannings_n is not None:
        corr = mannings_n_correction(Q, geometry, rough, geometry.mannings_n,
                                     solver="loglaw", tol=tol)

    if corr is not None:
        H, ustar, area = corr
    else:
        H, ustar, area = solve_depth_loglaw(Q, geometry, rough, tol=tol)

    if geometry.cross_section is not None and geometry.width is None:
        _, rh_val = _xs_area_rh(geometry.cross_section, H)
        ustar_phi = math.sqrt(C.g * rh_val * S)
    else:
        ustar_phi = ustar

    # ── transport calculation ─────────────────────────────────────────────────
    phi50 = ustar_phi ** 2 / (C.R * C.g * D50_m * taur50)

    qs_sum = 0.0
    p = []
    for i in range(n_frac):
        dmy = phi50 / (Di_m[i] / D50_m) ** beta
        gi  = g_func(dmy) * f[i]
        p.append(gi)
        qs_sum += gi

    if qs_sum > 0.0:
        p_norm = [pi / qs_sum for pi in p]
    else:
        p_norm = [0.0] * n_frac

    if geometry.width is not None and geometry.cross_section is None:
        qs_vol = qs_sum * geometry.width * ustar ** 3 / (C.R * C.g)
    else:
        qs_vol = qs_sum * ustar * S * area / C.R

    qs_kgs = qs_vol * C.rho_s
    fractions_kgs = [qs_kgs * p_norm[i] for i in range(n_frac)]

    return TransportResult(
        discharge_m3s=Q,
        total_bedload_kgs=qs_kgs,
        bedload_by_fraction=fractions_kgs,
        shields_stress=phi50,
        flow_depth_m=H,
        shear_velocity_ms=ustar,
    )


# ── D50-only (PKM82) variant ─────────────────────────────────────────────────

def transport_rate_d50(
    Q: float,
    geometry: ChannelGeometry,
    D50_mm: float,
    *,
    taur: float = C.PKM82_TAUR,
    dk: float   = C.PKM82_DK,
    tol: float  = 1e-5,
) -> TransportResult:
    """Parker-Klingeman-McLean (1982) single-fraction (D50-only) equation.

    Uses :func:`pkm_g` on the bulk Shields stress ``phi50``.
    No by-fraction breakdown is computed.
    """
    S = geometry.slope
    D50_m = D50_mm / 1000.0
    rough = dk * D50_m

    corr = None
    if geometry.mannings_n is not None:
        corr = mannings_n_correction(Q, geometry, rough, geometry.mannings_n,
                                     solver="loglaw", tol=tol)

    if corr is not None:
        H, ustar, area = corr
    else:
        H, ustar, area = solve_depth_loglaw(Q, geometry, rough, tol=tol)

    if geometry.cross_section is not None and geometry.width is None:
        _, rh_val = _xs_area_rh(geometry.cross_section, H)
        ustar_phi = math.sqrt(C.g * rh_val * S)
    else:
        ustar_phi = ustar

    phi50 = ustar_phi ** 2 / (C.R * C.g * D50_m * taur)
    w_star = pkm_g(phi50)

    if geometry.width is not None and geometry.cross_section is None:
        qs_vol = ustar ** 3 / (C.R * C.g) * geometry.width * w_star
    else:
        qs_vol = ustar * S * area / C.R * w_star

    qs_kgs = qs_vol * C.rho_s

    return TransportResult(
        discharge_m3s=Q,
        total_bedload_kgs=qs_kgs,
        bedload_by_fraction=[qs_kgs],
        shields_stress=phi50,
        flow_depth_m=H,
        shear_velocity_ms=ustar,
    )
