# BAGS-python

A pure-Python implementation of **BAGS** — *Bedload Assessment for Gravel-bed Streams* — a spreadsheet application developed by the U.S. Department of Agriculture Forest Service for calculating bedload sediment transport in gravel-bed rivers.

> **Source software:** The equations, constants, and calibration logic in this library are translated directly from the original BAGS VBA/Excel source code.
> Original BAGS software: **USDA Forest Service, Stream Systems Technology Center**
> Homepage: <http://www.stream.fs.fed.us/>
> Manual: Pitlick, J., Cui, Y., & Wilcock, P. (2009). *Manual for computing bed load transport using BAGS (Bedload Assessment for Gravel-bed Streams) software.* Gen. Tech. Rep. RMRS-GTR-223. Fort Collins, CO: USDA Forest Service, Rocky Mountain Research Station. <https://doi.org/10.2737/RMRS-GTR-223>

## Original authors

BAGS was designed and coded by **Dr. Yantao Cui**, with major technical contributions from:
John Pitlick · Peter Wilcock · John Potyondy · Paul Bakke · Yantao Cui

---

## What is BAGS?

BAGS implements six well-known bedload transport equations developed specifically for gravel-bed rivers. Given channel geometry, water-surface slope, bed-material grain-size distribution, and water discharge, each equation returns a predicted bedload transport rate.

| Module | Equation | Input GSD | Notes |
|---|---|---|---|
| `parker90` | Parker (1990) | Surface | Surface-based; uses Table A1 hiding correction |
| `parker82` | Parker & Klingeman (1982) | Substrate | Multi-fraction; PK or PKM G-function |
| `parker82` | Parker, Klingeman & McLean (1982) | Substrate D₅₀ | Single-fraction D₅₀ variant |
| `wilcock03` | Wilcock & Crowe (2003) | Surface | Surface-based mixed-size |
| `wilcock01` | Wilcock (2001) | Surface | Two-fraction (gravel + sand); requires calibration |
| `bakke99` | Bakke et al. (1999) | Substrate | Calibrated Parker-Klingeman variant |

## Disclaimer

> *"Use the software with your own judgement and at your own risk. Neither the Forest Service nor the authors are responsible for damages resulting from the application of this software."*
> — Original BAGS application agreement
 
---

## Installation

```bash
pip install .
```

Requires Python ≥ 3.10 and NumPy ≥ 1.22.

---

## Quick start

```python
from bags.data import GrainSizeDistribution, ChannelGeometry
from bags.equations import parker90

# Surface grain-size distribution: sizes (mm) and cumulative % finer
gsd = GrainSizeDistribution(
    sizes_mm  = [2,  4,  8,  16,  32,  64],
    finer_pct = [5, 15, 35,  65,  85, 100],
)

# Simple rectangular channel
geom = ChannelGeometry(slope=0.005, width=20.0)

result = parker90.transport_rate(Q=50.0, geometry=geom, surface_gsd=gsd)
print(f"{result.total_bedload_kgs:.2f} kg/s")
```

---

## Core data structures

### `GrainSizeDistribution`

```python
GrainSizeDistribution(
    sizes_mm,   # list[float] — grain diameters in mm, in ascending order
    finer_pct,  # list[float] — cumulative percent finer (0–100), same length as sizes_mm
)
```

| Parameter | Type | Description |
|---|---|---|
| `sizes_mm` | `list[float]` | Grain diameters in mm, strictly ascending. Minimum two values. Should span the full range of bed material present. |
| `finer_pct` | `list[float]` | Cumulative percent finer by weight (or volume) at each diameter. Values must be ascending, in the range 0–100. The final value should be 100. |

**Notes:**
- The GSD is interpolated using log-linear interpolation in size (geometric interpolation), consistent with standard sieve analysis conventions.
- For Parker (1990) and Wilcock & Crowe (2003), supply the **surface** (pavement) GSD. For Parker-Klingeman (1982) and Bakke et al. (1999), supply the **substrate** (sub-surface) GSD.
- The library internally computes D₅₀, D₆₅, D₉₀, geometric mean Dsg, and per-fraction volumes from this distribution.

**Example:**
```python
# Gravel-bed stream: 5 % finer than 2 mm, 100 % finer than 64 mm
gsd = GrainSizeDistribution(
    sizes_mm  = [2,  4,  8, 16, 32,  64],
    finer_pct = [5, 15, 35, 65, 85, 100],
)
```

---

### `ChannelGeometry`

```python
ChannelGeometry(
    slope,                # float  — water-surface slope (m/m), required
    width=None,           # float  — bankfull width in m
    cross_section=None,   # list[tuple[float, float]] — surveyed cross-section
    mannings_n=None,      # float  — main-channel Manning's n
    mannings_n_left=None, # float  — left-floodplain Manning's n
    mannings_n_right=None,# float  — right-floodplain Manning's n
)
```

| Parameter | Type | Default | Description |
|---|---|---|---|
| `slope` | `float` | required | Water-surface (energy) slope, dimensionless (m/m). Typical gravel-bed range: 0.001–0.05. |
| `width` | `float \| None` | `None` | Bankfull channel width in metres. Use for simple rectangular cross-sections. Mutually exclusive with `cross_section` — supply exactly one. |
| `cross_section` | `list[tuple[float, float]] \| None` | `None` | Surveyed cross-section as a list of `(station_m, elevation_m)` pairs in left-to-right order. The wetted area and hydraulic radius at each trial depth are computed by panel integration. Mutually exclusive with `width`. |
| `mannings_n` | `float \| None` | `None` | Main-channel Manning's roughness coefficient. When provided and greater than the grain-roughness equivalent, the hydraulic depth is adjusted downward to account for form drag. Typical values: 0.025–0.060 for natural gravel-bed channels. |
| `mannings_n_left` | `float \| None` | `None` | Manning's n for the left floodplain (used only with `cross_section`). |
| `mannings_n_right` | `float \| None` | `None` | Manning's n for the right floodplain (used only with `cross_section`). |

**Notes:**
- Exactly one of `width` or `cross_section` must be supplied.
- Slope is used for computing shear velocity: u* = √(g · H · S) for rectangular channels, u* = √(g · Rh · S) for cross-sections.
- When `mannings_n` is supplied, the hydraulic depth is iterated so that the Manning discharge matches the grain-roughness log-law discharge. This partitions total shear stress into grain and form-drag components, leaving only the grain-related portion for transport calculations.

**Examples:**
```python
# Rectangular channel, no roughness correction
geom = ChannelGeometry(slope=0.005, width=20.0)

# With Manning's n roughness correction
geom = ChannelGeometry(slope=0.005, width=20.0, mannings_n=0.035)

# Surveyed cross-section
geom = ChannelGeometry(
    slope=0.003,
    cross_section=[
        (0.0, 102.5), (5.0, 100.0), (10.0, 99.2),
        (20.0, 99.0), (30.0, 99.3), (35.0, 100.1), (40.0, 102.8),
    ],
)
```

---

### `TransportResult`

All equation functions return a `TransportResult`:

```python
TransportResult(
    discharge_m3s,        # float       — input discharge (m³/s)
    total_bedload_kgs,    # float       — total bedload transport rate (kg/s)
    bedload_by_fraction,  # list[float] — transport rate per size class (kg/s)
    shields_stress,       # float       — normalised Shields stress (equation-specific)
    flow_depth_m,         # float       — computed flow depth (m)
    shear_velocity_ms,    # float       — bed shear velocity u* (m/s)
)
```

| Field | Units | Description |
|---|---|---|
| `discharge_m3s` | m³/s | Echo of the input discharge Q. |
| `total_bedload_kgs` | kg/s | Total bedload transport rate summed across all size fractions. |
| `bedload_by_fraction` | kg/s | Transport rate for each size class, in the same order as the input GSD fractions. For Wilcock (2001) this is `[gravel_kgs, sand_kgs]`. |
| `shields_stress` | — | Normalised Shields stress. Definition varies by equation: φ*sgo for Parker (1990), φ*₅₀ for PK82/Bakke99, τ/τ*rsg for Wilcock & Crowe (2003), 0.0 for Wilcock (2001). |
| `flow_depth_m` | m | Maximum flow depth solved from the discharge–depth relationship. For cross-sections this is the depth above the thalweg. |
| `shear_velocity_ms` | m/s | Bed shear velocity u* = √(g · H · S) or √(g · Rh · S) for cross-sections. |

---

## Equations

### Parker (1990)

Surface-based equation with grain-size-dependent hiding correction via Table A1 lookup. Roughness is computed as `rough = Dk × D₉₀`. Uses log-law hydraulics.

```python
from bags.equations import parker90

result = parker90.transport_rate(Q, geometry, surface_gsd)
```

**Required parameters:**

| Parameter | Type | Description |
|---|---|---|
| `Q` | `float` | Water discharge (m³/s). |
| `geometry` | `ChannelGeometry` | Channel geometry. Supply `width` or `cross_section`. |
| `surface_gsd` | `GrainSizeDistribution` | Surface (pavement) grain-size distribution. Used to compute geometric mean Dsg, standard deviation σ_φ, and D₉₀. Must span at least the gravel range (≥ 2 mm). |

**Optional keyword parameters (scientific defaults):**

| Parameter | Default | Description |
|---|---|---|
| `taursgo` | `0.0386` | Reference Shields stress τ*sgo for the geometric mean grain size. |
| `alpha` | `0.00218` | Bedload efficiency coefficient. |
| `beta` | `0.0951` | Hiding-function exponent controlling size-selective transport. |
| `dk` | `2.0` | Roughness multiplier: `rough = dk × D₉₀`. |
| `tol` | `1e-5` | Convergence tolerance for the depth solver. |

---

### Parker & Klingeman (1982) — multi-fraction

Substrate-based equation using the Parker-Klingeman (PK) or Parker-Klingeman-McLean (PKM) G-function. Roughness is `rough = Dk × D₅₀`. Uses log-law hydraulics.

```python
from bags.equations import parker82

result = parker82.transport_rate(Q, geometry, substrate_gsd)
result = parker82.transport_rate(Q, geometry, substrate_gsd, use_pkm=True)
```

**Required parameters:**

| Parameter | Type | Description |
|---|---|---|
| `Q` | `float` | Water discharge (m³/s). |
| `geometry` | `ChannelGeometry` | Channel geometry. Supply `width` or `cross_section`. |
| `substrate_gsd` | `GrainSizeDistribution` | Substrate (sub-surface) grain-size distribution. Used to compute D₅₀ (for roughness and the reference Shields stress) and the volume fraction in each size class. |

**Optional keyword parameters (scientific defaults):**

| Parameter | Default | Description |
|---|---|---|
| `use_pkm` | `False` | If `True`, uses the Parker-Klingeman-McLean W* transport function (three-branch piecewise) instead of the PK two-branch function. |
| `taur50` | `0.0876` | Reference Shields stress based on substrate D₅₀. |
| `beta` | `0.018` | Hiding-function exponent: controls how strongly coarser grains are hidden by finer ones. |
| `dk` | `10.7` | Roughness multiplier: `rough = dk × D₅₀`. |
| `tol` | `1e-5` | Convergence tolerance for the depth solver. |

---

### Parker, Klingeman & McLean (1982) — D₅₀ variant

Single-fraction version of the PKM equation. Requires only the substrate D₅₀; no full GSD is needed. Does not produce a by-fraction breakdown.

```python
from bags.equations import parker82

result = parker82.transport_rate_d50(Q, geometry, D50_mm=16.0)
```

**Required parameters:**

| Parameter | Type | Description |
|---|---|---|
| `Q` | `float` | Water discharge (m³/s). |
| `geometry` | `ChannelGeometry` | Channel geometry. Supply `width` or `cross_section`. |
| `D50_mm` | `float` | Substrate median grain diameter D₅₀ in millimetres. Used to compute roughness and the bulk Shields stress. |

**Optional keyword parameters (scientific defaults):**

| Parameter | Default | Description |
|---|---|---|
| `taur` | `0.0876` | Reference Shields stress based on D₅₀. |
| `dk` | `10.7` | Roughness multiplier: `rough = dk × D₅₀`. |
| `tol` | `1e-5` | Convergence tolerance for the depth solver. |

---

### Wilcock & Crowe (2003)

Surface-based mixed-size equation. Uses Manning-Strickler hydraulics with roughness `rough = 2 × D₆₅`. The reference stress and hiding/exposure correction are computed from the surface GSD.

```python
from bags.equations import wilcock03

result = wilcock03.transport_rate(Q, geometry, surface_gsd)
```

**Required parameters:**

| Parameter | Type | Description |
|---|---|---|
| `Q` | `float` | Water discharge (m³/s). |
| `geometry` | `ChannelGeometry` | Channel geometry. Supply `width` or `cross_section`. |
| `surface_gsd` | `GrainSizeDistribution` | Surface (pavement) grain-size distribution. Used to compute geometric mean Dsg, D₆₅ (for roughness), sand fraction Fs, and volume fractions per size class. The sand fraction Fs (proportion finer than 2 mm) is critical — it controls the reference Shields stress through the equation τ*rsg = ρRgDsg(0.021 + 0.015/exp(20·Fs)). |

**Optional keyword parameters (scientific defaults):**

| Parameter | Default | Description |
|---|---|---|
| `dk` | `2.0` | Roughness multiplier: `rough = dk × D₆₅`. |
| `tol` | `1e-5` | Convergence tolerance for the depth solver. |

**Note on `shields_stress` in result:** Returns τ / τ*rsg (transport stage normalised to the geometric-mean reference stress), not a per-fraction value.

---

### Wilcock (2001) — two-fraction calibrated

Separates bedload into a gravel fraction and a sand fraction, each with its own reference shear stress. Requires calibration against field observations of gravel and sand transport before use. Uses Manning-Strickler hydraulics with roughness `rough = 2 × D₆₅`.

#### Step 1: Calibrate

```python
from bags.equations import wilcock01

TaurG, TaurS = wilcock01.calibrate(
    discharges,
    observed_total_kgs,
    observed_gravel_fraction,
    geometry,
    surface_gsd,
)
```

**Required parameters:**

| Parameter | Type | Description |
|---|---|---|
| `discharges` | `Sequence[float]` | Measured water discharges (m³/s), one value per bedload sample. |
| `observed_total_kgs` | `Sequence[float]` | Total measured bedload transport rate (kg/s) for each sample. Values ≤ 0 are skipped. |
| `observed_gravel_fraction` | `Sequence[float]` | Fraction of total bedload that is gravel (coarser than 2 mm), as a proportion from 0 to 1, for each sample. Used to partition total bedload into `observed_gravel_kgs` and `observed_sand_kgs`. |
| `geometry` | `ChannelGeometry` | Channel geometry. Must match the geometry used for prediction. |
| `surface_gsd` | `GrainSizeDistribution` | Surface grain-size distribution. Used for D₆₅ (roughness) and to compute the bulk gravel fraction Fg and sand fraction Fs. |

**Optional keyword parameters:**

| Parameter | Default | Description |
|---|---|---|
| `dk` | `2.0` | Roughness multiplier: `rough = dk × D₆₅`. |
| `tol` | `1e-5` | Convergence tolerance for the depth solver. |

**Returns:** `(TaurG, TaurS)` — calibrated reference shear stresses in Pa for gravel and sand respectively. Pass these directly to `transport_rate`.

#### Step 2: Predict

```python
result = wilcock01.transport_rate(Q, geometry, surface_gsd, TaurG, TaurS)
```

**Required parameters:**

| Parameter | Type | Description |
|---|---|---|
| `Q` | `float` | Water discharge (m³/s). |
| `geometry` | `ChannelGeometry` | Channel geometry. Must match that used for calibration. |
| `surface_gsd` | `GrainSizeDistribution` | Surface grain-size distribution (same as used for calibration). |
| `TaurG` | `float` | Calibrated reference shear stress for gravel (Pa), from `calibrate()`. |
| `TaurS` | `float` | Calibrated reference shear stress for sand (Pa), from `calibrate()`. |

**Optional keyword parameters:**

| Parameter | Default | Description |
|---|---|---|
| `dk` | `2.0` | Roughness multiplier: `rough = dk × D₆₅`. Must match the value used during calibration. |
| `tol` | `1e-5` | Convergence tolerance for the depth solver. |

**Note on result:** `bedload_by_fraction` = `[gravel_kgs, sand_kgs]`. `shields_stress` is 0.0 (no single normalising stress is defined for the two-fraction model).

---

### Bakke et al. (1999) — calibrated PK

Iteratively calibrates the PK82 reference Shields stress (τ*₅₀) and hiding exponent (β) to observed bedload measurements. Optionally uses observed bedload D₅₀ to improve β calibration, matching the VBA `GoGetExponent` step. After calibration, transport is computed with the PK or PKM G-function.

#### Step 1: Calibrate

```python
from bags.equations import bakke99

taur50, beta = bakke99.calibrate(
    discharges,
    observed_kgs,
    geometry,
    substrate_gsd,
    observed_d50_mm=observed_d50_mm,   # optional; improves β calibration
)
```

**Required parameters:**

| Parameter | Type | Description |
|---|---|---|
| `discharges` | `Sequence[float]` | Measured water discharges (m³/s), one per bedload sample. |
| `observed_kgs` | `Sequence[float]` | Total measured bedload transport rate (kg/s) for each sample. Values ≤ 0 are skipped. |
| `geometry` | `ChannelGeometry` | Channel geometry. Must match the geometry used for prediction. |
| `substrate_gsd` | `GrainSizeDistribution` | Substrate (sub-surface) grain-size distribution. Used for D₅₀ (roughness and reference stress) and the volume fraction in each size class. |

**Optional keyword parameters:**

| Parameter | Default | Description |
|---|---|---|
| `observed_d50_mm` | `None` | Observed bedload D₅₀ in mm for each sample (same length as `discharges`). When provided, β is calibrated by minimising squared log residuals of predicted vs. observed bedload D₅₀ — this matches the original VBA `GoGetExponent` routine and is the physically preferred approach since β controls size-selectivity, not total transport magnitude. When `None`, β is calibrated against total bedload instead. |
| `use_pkm` | `False` | If `True`, use the PKM G-function (three-branch) during calibration and transport. |
| `dk` | `10.7` | Roughness multiplier: `rough = dk × D₅₀`. |
| `max_iterations` | `6` | Maximum number of alternating optimisation cycles. Each cycle first optimises τ*₅₀ (holding β fixed), then optimises β (holding τ*₅₀ fixed). Convergence is declared when both parameters change by less than 0.1 % and at least 3 cycles have completed. |
| `tol` | `1e-5` | Convergence tolerance for the depth solver. |

**Returns:** `(taur50, beta)` — calibrated reference Shields stress and hiding exponent. Pass these directly to `transport_rate`.

#### Step 2: Predict

```python
result = bakke99.transport_rate(Q, geometry, substrate_gsd, taur50, beta)
```

**Required parameters:**

| Parameter | Type | Description |
|---|---|---|
| `Q` | `float` | Water discharge (m³/s). |
| `geometry` | `ChannelGeometry` | Channel geometry. Must match that used for calibration. |
| `substrate_gsd` | `GrainSizeDistribution` | Substrate grain-size distribution (same as used for calibration). |
| `taur50` | `float` | Calibrated reference Shields stress, from `calibrate()`. |
| `beta` | `float` | Calibrated hiding-function exponent, from `calibrate()`. |

**Optional keyword parameters:**

| Parameter | Default | Description |
|---|---|---|
| `use_pkm` | `False` | Use PKM G-function. Must match the value used during calibration. |
| `dk` | `10.7` | Roughness multiplier. Must match the value used during calibration. |
| `tol` | `1e-5` | Convergence tolerance for the depth solver. |

---

## Loading data from files

```python
from bags.io import load_gsd, load_cross_section, load_geometry

gsd  = load_gsd("surface_gsd.csv")            # two-column CSV: size_mm, pct_finer
xs   = load_cross_section("channel_xs.csv")   # two-column CSV: station_m, elev_m
geom = load_geometry("channel.json")          # JSON with slope, width, cross_section, …
```

## Monte Carlo / uncertainty analysis

All equation constants are keyword arguments with scientific defaults, enabling perturbation for Monte Carlo uncertainty analysis:

```python
import numpy as np

results = [
    parker90.transport_rate(
        Q=50.0, geometry=geom, surface_gsd=gsd,
        taursgo=np.random.normal(0.0386, 0.003),
    )
    for _ in range(1000)
]
```

---

## References

- Parker, G. (1990). Surface-based bedload transport relation for gravel rivers. *Journal of Hydraulic Research*, 28(4), 417–436.
- Parker, G., & Klingeman, P. C. (1982). On why gravel bed streams are paved. *Water Resources Research*, 18(5), 1409–1423.
- Parker, G., Klingeman, P. C., & McLean, D. G. (1982). Bedload and size distribution in paved gravel-bed streams. *Journal of the Hydraulics Division*, 108(HY4), 544–571.
- Wilcock, P. R. (2001). Toward a practical method for estimating sediment-transport rates in gravel-bed rivers. *Earth Surface Processes and Landforms*, 26(13), 1395–1408.
- Wilcock, P. R., & Crowe, J. C. (2003). Surface-based transport model for mixed-size sediment. *Journal of Hydraulic Engineering*, 129(2), 120–128.
- Bakke, P. D., Basdekas, P. O., Dawdy, D. R., & Klingeman, P. C. (1999). Calibrated Parker-Klingeman model for gravel transport. *Journal of Hydraulic Engineering*, 125(6), 657–660.
- Pitlick, J., Cui, Y., & Wilcock, P. (2009). *Manual for computing bed load transport using BAGS software.* RMRS-GTR-223. USDA Forest Service. <https://doi.org/10.2737/RMRS-GTR-223>
