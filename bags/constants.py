"""Physical constants and look-up tables used by the BAGS equations."""
import numpy as np

# ── Physical constants ────────────────────────────────────────────────────────
g: float = 9.81       # gravitational acceleration (m/s²)
R: float = 1.65       # submerged specific gravity of sediment (ρ_s/ρ_w − 1)
rho: float = 1000.0   # water density (kg/m³)
rho_s: float = 2650.0 # sediment density (kg/m³)

# ── Parker (1990) equation defaults ──────────────────────────────────────────
PARKER90_TAURSGO: float = 0.0386   # reference Shields stress
PARKER90_ALPHA: float  = 0.00218   # bedload coefficient
PARKER90_BETA: float   = 0.0951    # hiding-function exponent
PARKER90_DK: float     = 2.0       # roughness coefficient (rough = Dk * D90)

# ── Parker-Klingeman (1982) / PKM defaults ────────────────────────────────────
PK82_TAUR50: float  = 0.0876   # reference Shields stress (D50-based)
PK82_BETA: float    = 0.018    # hiding-function exponent
PK82_DK: float      = 10.7    # roughness coefficient (rough = Dk * D50)

# ── Parker-Klingeman-McLean (1982) D50 variant ────────────────────────────────
PKM82_TAUR: float   = 0.0876
PKM82_DK: float     = 10.7

# ── Wilcock (2001) two-fraction / Wilcock & Crowe (2003) ─────────────────────
WILCOCK_DK: float   = 2.0    # roughness = 2 × D65

# ── Parker (1990) Table A1 ────────────────────────────────────────────────────
# 36-row look-up: phi_sgo*, omega_0, sigma_0
# Source: Parker (1990) J. Hydraulic Research 28(4), Appendix A
PARKER90_PHI: np.ndarray = np.array([
    0.31, 0.34, 0.38, 0.42, 0.46, 0.51, 0.56, 0.62, 0.68, 0.75,
    0.83, 0.91, 1.00, 1.10, 1.21, 1.33, 1.47, 1.61, 1.77, 1.95,
    2.15, 2.36, 2.59, 2.85, 3.13, 3.45, 3.79, 4.17, 4.58, 5.04,
    5.54, 6.10, 6.71, 7.38, 8.11, 8.92,
], dtype=float)

PARKER90_OMEGA0: np.ndarray = np.array([
    1.000, 1.000, 1.000, 1.000, 1.000, 1.000, 1.000, 1.000, 1.000, 1.000,
    1.000, 1.000, 1.000, 1.030, 1.065, 1.100, 1.130, 1.153, 1.175, 1.190,
    1.202, 1.212, 1.220, 1.226, 1.232, 1.236, 1.240, 1.243, 1.246, 1.248,
    1.250, 1.252, 1.253, 1.254, 1.255, 1.256,
], dtype=float)

PARKER90_SIGMA0: np.ndarray = np.array([
    1.000, 1.000, 1.000, 1.000, 1.000, 1.000, 1.000, 1.000, 1.000, 1.000,
    1.000, 1.000, 1.000, 0.967, 0.934, 0.902, 0.873, 0.845, 0.821, 0.799,
    0.779, 0.761, 0.744, 0.729, 0.715, 0.703, 0.691, 0.680, 0.671, 0.663,
    0.655, 0.648, 0.642, 0.636, 0.631, 0.626,
], dtype=float)
