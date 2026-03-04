"""Parker (1990) surface-based bedload transport equation.

Reference
---------
Parker, G. (1990). Surface-based bedload transport relation for gravel rivers.
    *Journal of Hydraulic Research*, 28(4), 417-436.

Notes
-----
* Roughness length: ``rough = Dk × D90``  (default Dk = 2.0)
* Depth solver: log-law bisection (or Manning's n correction if supplied)
* Transport parameter: volumetric bedload (m³/s), converted to kg/s via ρ_s.
"""
from __future__ import annotations
import math
import numpy as np

from ..data import GrainSizeDistribution, ChannelGeometry, TransportResult
from .. import constants as C
from ..grain_size import geometric_mean_phi, percentile, fraction_volumes
from ..hydraulics import (
    solve_depth_loglaw, _xs_area_rh, mannings_n_correction,
)


# ── G function (Parker 1990 eq. A-1) ─────────────────────────────────────────

def parker_g(xx: float) -> float:
    """Piecewise G function from Parker (1990)."""
    if xx <= 1.0:
        return xx ** 14.2
    if xx <= 1.59:
        return math.exp(14.2 * (xx - 1.0) - 9.28 * (xx - 1.0) ** 2)
    return 5474.0 * (1.0 - 0.853 / xx) ** 4.5


# ── Table A1 interpolation ────────────────────────────────────────────────────

def omega_sigma_interp(phisgo: float) -> tuple[float, float]:
    """Interpolate (omega_0, sigma_0) from Parker Table A1."""
    phi_arr = C.PARKER90_PHI
    om_arr  = C.PARKER90_OMEGA0
    sg_arr  = C.PARKER90_SIGMA0

    omega0 = float(np.interp(phisgo, phi_arr, om_arr))
    sigma0 = float(np.interp(phisgo, phi_arr, sg_arr))
    return omega0, sigma0


# ── main transport function ───────────────────────────────────────────────────

def transport_rate(
    Q: float,
    geometry: ChannelGeometry,
    surface_gsd: GrainSizeDistribution,
    *,
    taursgo: float = C.PARKER90_TAURSGO,
    alpha: float   = C.PARKER90_ALPHA,
    beta: float    = C.PARKER90_BETA,
    dk: float      = C.PARKER90_DK,
    tol: float     = 1e-5,
) -> TransportResult:
    """Compute bedload transport with the Parker (1990) surface-based equation.

    Parameters
    ----------
    Q : float
        Water discharge (m³/s).
    geometry : ChannelGeometry
        Channel geometry.  Supply either ``width`` or ``cross_section``.
    surface_gsd : GrainSizeDistribution
        Surface grain-size distribution.
    taursgo : float
        Reference Shields stress (default 0.0386).
    alpha : float
        Bedload coefficient (default 0.00218).
    beta : float
        Hiding-function exponent (default 0.0951).
    dk : float
        Roughness multiplier: ``rough = dk × D90``  (default 2.0).

    Returns
    -------
    TransportResult
        Bedload transport result; ``total_bedload_kgs`` in kg/s.
    """
    S = geometry.slope

    # grain-size statistics
    Dsg_m, std_phi = geometric_mean_phi(surface_gsd)
    D90_mm = percentile(surface_gsd, 90.0)
    D90_m  = D90_mm / 1000.0
    rough  = dk * D90_m

    # size fractions
    psi, Di_m, f = fraction_volumes(surface_gsd)
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

    # cross-section: compute Rh for phisgo
    if geometry.cross_section is not None and geometry.width is None:
        _, rh_val = _xs_area_rh(geometry.cross_section, H)
        ustar_phi = math.sqrt(C.g * rh_val * S)
    else:
        ustar_phi = ustar

    # ── transport calculation ─────────────────────────────────────────────────
    phisgo = ustar_phi ** 2 / (C.R * C.g * Dsg_m * taursgo)
    omega0, sigma0 = omega_sigma_interp(phisgo)
    omega = 1.0 + std_phi / sigma0 * (omega0 - 1.0)

    qs_sum = 0.0
    p = []
    for i in range(n_frac):
        xx = omega * phisgo * (Dsg_m / Di_m[i]) ** beta
        gi = parker_g(xx) * f[i]
        p.append(gi)
        qs_sum += gi

    # normalised bedload fractions
    if qs_sum > 0.0:
        p_norm = [pi / qs_sum for pi in p]
    else:
        p_norm = [0.0] * n_frac

    # total transport (m³/s)
    if geometry.width is not None and geometry.cross_section is None:
        qs_vol = alpha * ustar ** 3 / (C.R * C.g) * geometry.width * qs_sum
    else:
        qs_vol = alpha * ustar * S * area / C.R * qs_sum

    qs_kgs = qs_vol * C.rho_s
    fractions_kgs = [qs_kgs * p_norm[i] for i in range(n_frac)]

    return TransportResult(
        discharge_m3s=Q,
        total_bedload_kgs=qs_kgs,
        bedload_by_fraction=fractions_kgs,
        shields_stress=phisgo,
        flow_depth_m=H,
        shear_velocity_ms=ustar,
    )
