"""Hydraulic solvers used by the BAGS transport equations.

Two roughness/resistance frameworks are supported:

* **Log-law** (Parker 1990, Parker-Klingeman 1982):
  ``Q = 2.5 · W · H · √(g·H·S) · ln(11·H / rough)``

* **Manning-Strickler** (Wilcock 2001/2003):
  ``Q = 8.1 · W · H · √(g·H·S) · (H/rough)^(1/6)``
  which has an explicit solution for constant-width channels.
"""
from __future__ import annotations
import math
from .data import ChannelGeometry
from . import constants as C


# ── cross-section geometry helpers ───────────────────────────────────────────

def _xs_area_rh(cross_section: list[tuple[float, float]], H: float
                ) -> tuple[float, float]:
    """Area (m²) and hydraulic radius (m) for the full cross-section at depth *H*.

    *H* is measured above the thalweg (minimum elevation in the cross-section).
    """
    z_min = min(z for _, z in cross_section)
    z_ws = z_min + H

    area = 0.0
    perimeter = 0.0

    for i in range(len(cross_section) - 1):
        x1, z1 = cross_section[i]
        x2, z2 = cross_section[i + 1]
        dx = abs(x2 - x1)
        dz = z2 - z1
        panel_len = math.sqrt(dx * dx + dz * dz)

        if z1 >= z_ws and z2 >= z_ws:
            # both above water surface — no contribution
            continue

        if z1 <= z_ws and z2 <= z_ws:
            # both submerged — full trapezoidal panel
            d1 = z_ws - z1
            d2 = z_ws - z2
            area += 0.5 * (d1 + d2) * dx
            perimeter += panel_len
        elif z1 < z_ws <= z2:
            # left submerged, right above — triangle
            frac = (z_ws - z1) / (z2 - z1)
            x_int = x1 + frac * (x2 - x1)
            area += 0.5 * (z_ws - z1) * abs(x_int - x1)
            perimeter += frac * panel_len
        else:
            # right submerged, left above — triangle
            frac = (z_ws - z2) / (z1 - z2)
            x_int = x2 + frac * (x1 - x2)
            area += 0.5 * (z_ws - z2) * abs(x_int - x2)
            perimeter += frac * panel_len

    rh = area / perimeter if perimeter > 1e-12 else 0.0
    return area, rh


# ── discharge functions ───────────────────────────────────────────────────────

def _q_loglaw_width(H: float, W: float, S: float, rough: float) -> float:
    """Discharge for constant-width log-law channel."""
    if H <= 0.0 or rough <= 0.0:
        return 0.0
    ratio = 11.0 * H / rough
    if ratio <= 1.0:
        return 0.0
    return 2.5 * W * H * math.sqrt(C.g * H * S) * math.log(ratio)


def _q_loglaw_xs(H: float, xs: list[tuple[float, float]], S: float,
                 rough: float,
                 mn_left: float | None = None,
                 mn_right: float | None = None,
                 ) -> float:
    """Discharge for cross-section using log-law for the main channel.

    Floodplains use Manning's n if provided (simplified: treated as separate
    trapezoidal panels beyond the channel banks — currently not implemented
    separately; both floodplain Manning's n values are ignored and the full
    cross-section is used as the main channel).
    """
    area, rh = _xs_area_rh(xs, H)
    if rh <= 0.0 or area <= 0.0:
        return 0.0
    ratio = 11.0 * rh / rough
    if ratio <= 1.0:
        return 0.0
    return area * math.sqrt(C.g * rh * S) * 2.5 * math.log(ratio)


def _q_strickler_width(H: float, W: float, S: float, rough: float) -> float:
    """Discharge for constant-width Manning-Strickler channel."""
    if H <= 0.0 or rough <= 0.0:
        return 0.0
    return 8.1 * W * H * math.sqrt(C.g * H * S) * (H / rough) ** (1.0 / 6.0)


def _q_strickler_xs(H: float, xs: list[tuple[float, float]], S: float,
                    rough: float,
                    mn_left: float | None = None,
                    mn_right: float | None = None,
                    ) -> float:
    """Discharge for cross-section using Manning-Strickler for main channel."""
    area, rh = _xs_area_rh(xs, H)
    if rh <= 0.0 or area <= 0.0:
        return 0.0
    q_main = 8.1 * area * math.sqrt(C.g * rh * S) * (rh / rough) ** (1.0 / 6.0)
    return q_main


# ── solvers ───────────────────────────────────────────────────────────────────

def _bisect_depth(q_func, Q_target: float, H_lo: float = 1e-6,
                  H_hi: float | None = None, tol: float = 1e-5,
                  max_iter: int = 200) -> float:
    """Return depth H such that q_func(H) ≈ Q_target via bisection."""
    if H_hi is None:
        # expand upper bracket until Q(H_hi) >= Q_target
        H_hi = max(1.0, H_lo * 1000)
        for _ in range(80):
            if q_func(H_hi) >= Q_target:
                break
            H_hi *= 2.0

    for _ in range(max_iter):
        H_mid = 0.5 * (H_lo + H_hi)
        q_mid = q_func(H_mid)
        if abs(q_mid - Q_target) / max(Q_target, 1e-15) < tol:
            return H_mid
        if q_mid < Q_target:
            H_lo = H_mid
        else:
            H_hi = H_mid

    return 0.5 * (H_lo + H_hi)


def solve_depth_loglaw(Q: float, geom: ChannelGeometry, rough: float,
                       tol: float = 1e-5) -> tuple[float, float, float]:
    """Solve for flow depth using the log-law velocity profile.

    Returns (H, ustar, area):
      * H     — flow depth (or max depth for cross-section), m
      * ustar — shear velocity = sqrt(g·Rh·S), m/s
      * area  — cross-sectional flow area, m²
    """
    S = geom.slope

    if geom.width is not None and geom.cross_section is None:
        W = geom.width
        q_func = lambda H: _q_loglaw_width(H, W, S, rough)
        H = _bisect_depth(q_func, Q, tol=tol)
        ustar = math.sqrt(C.g * H * S)
        area = W * H
    else:
        xs = geom.cross_section
        q_func = lambda H: _q_loglaw_xs(H, xs, S, rough,
                                         geom.mannings_n_left,
                                         geom.mannings_n_right)
        H_hi = (max(z for _, z in xs) - min(z for _, z in xs)) * 3.0
        H = _bisect_depth(q_func, Q, H_hi=H_hi, tol=tol)
        area, rh = _xs_area_rh(xs, H)
        ustar = math.sqrt(C.g * rh * S)

    return H, ustar, area


def solve_depth_strickler(Q: float, geom: ChannelGeometry, rough: float,
                           tol: float = 1e-5) -> tuple[float, float, float]:
    """Solve for flow depth using the Manning-Strickler formula.

    For constant-width channels uses the explicit formula.
    Returns (H, ustar, area).
    """
    S = geom.slope

    if geom.width is not None and geom.cross_section is None:
        W = geom.width
        # explicit solution: H = (Q / 8.1 / W)^0.6 * rough^0.1 / (g*S)^0.3
        H = (Q / 8.1 / W) ** 0.6 * rough ** 0.1 / (C.g * S) ** 0.3
        ustar = math.sqrt(C.g * H * S)
        area = W * H
    else:
        xs = geom.cross_section
        q_func = lambda H: _q_strickler_xs(H, xs, S, rough,
                                            geom.mannings_n_left,
                                            geom.mannings_n_right)
        H_hi = (max(z for _, z in xs) - min(z for _, z in xs)) * 3.0
        H = _bisect_depth(q_func, Q, H_hi=H_hi, tol=tol)
        area, rh = _xs_area_rh(xs, H)
        ustar = math.sqrt(C.g * rh * S)

    return H, ustar, area


# ── Manning's n roughness correction ─────────────────────────────────────────

def mannings_n_correction(
    Q: float, geom: ChannelGeometry, rough: float,
    mannings_n: float, solver: str = "loglaw",
    tol: float = 1e-5,
) -> tuple[float, float, float] | None:
    """Apply Manning's n roughness correction (2006 BAGS revision).

    Returns (H, ustar, area) with grain-roughness correction applied, or
    *None* if Manning's n is smaller than grain roughness (correction unused).

    *solver* is ``"loglaw"`` (Parker 90/82) or ``"strickler"`` (Wilcock).
    """
    nD = 0.04 * rough ** (1.0 / 6.0)
    if mannings_n <= nD:
        return None  # grain roughness dominates; no correction needed

    # depth from Manning's n
    if geom.width is not None and geom.cross_section is None:
        H = (mannings_n * Q / geom.width / math.sqrt(geom.slope)) ** (3.0 / 5.0)
        tau_total = C.rho * C.g * H * geom.slope
        area = geom.width * H
    else:
        # cross-section: bisect on Manning's n for main channel
        xs = geom.cross_section
        S = geom.slope
        q_func = lambda H: (lambda a, r: a * r ** (2.0 / 3.0) * S ** 0.5 / mannings_n)(
            *_xs_area_rh(xs, H))
        H_hi = (max(z for _, z in xs) - min(z for _, z in xs)) * 3.0
        H = _bisect_depth(q_func, Q, H_hi=H_hi, tol=tol)
        area, rh = _xs_area_rh(xs, H)
        tau_total = C.rho * C.g * rh * S

    tau_grain = tau_total * (nD / mannings_n) ** 1.5
    ustar = math.sqrt(tau_grain / C.rho)
    return H, ustar, area
