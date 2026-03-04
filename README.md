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

```python
GrainSizeDistribution(sizes_mm, finer_pct)

ChannelGeometry(
    slope,               # dimensionless (m/m)
    width=None,          # bankfull width in m  ─┐ supply one
    cross_section=None,  # [(station_m, elev_m)] ─┘
    mannings_n=None,     # optional roughness correction
)

# All equations return:
TransportResult(
    discharge_m3s,
    total_bedload_kgs,       # kg/s
    bedload_by_fraction,     # kg/s per size class
    shields_stress,
    flow_depth_m,
    shear_velocity_ms,
)
```

---

## Equations

### Parker (1990)

Surface-based equation with grain-size-dependent hiding correction via Table A1 lookup.

```python
from bags.equations import parker90

result = parker90.transport_rate(Q, geometry, surface_gsd)
```

### Parker & Klingeman (1982) — multi-fraction

Substrate-based equation. The `use_pkm=True` flag selects the Parker-Klingeman-McLean G-function.

```python
from bags.equations import parker82

result = parker82.transport_rate(Q, geometry, substrate_gsd)
result = parker82.transport_rate(Q, geometry, substrate_gsd, use_pkm=True)
```

### Parker, Klingeman & McLean (1982) — D₅₀ variant

Single-fraction version using only the substrate D₅₀.

```python
result = parker82.transport_rate_d50(Q, geometry, D50_mm=16.0)
```

### Wilcock & Crowe (2003)

Surface-based mixed-size equation using Manning-Strickler hydraulics.

```python
from bags.equations import wilcock03

result = wilcock03.transport_rate(Q, geometry, surface_gsd)
```

### Wilcock (2001) — two-fraction calibrated

Requires calibration against observed gravel and sand bedload samples.

```python
from bags.equations import wilcock01

TaurG, TaurS = wilcock01.calibrate(
    discharges, observed_total_kgs, observed_gravel_fraction,
    geometry, surface_gsd,
)
result = wilcock01.transport_rate(Q, geometry, surface_gsd, TaurG, TaurS)
```

### Bakke et al. (1999) — calibrated PK

Iteratively calibrates the PK82 reference Shields stress (τ*₅₀) and hiding exponent (β) to observed total bedload and (optionally) bedload D₅₀.

```python
from bags.equations import bakke99

taur50, beta = bakke99.calibrate(
    discharges, observed_kgs, geometry, substrate_gsd,
    observed_d50_mm=observed_d50_mm,   # optional; improves β calibration
)
result = bakke99.transport_rate(Q, geometry, substrate_gsd, taur50, beta)
```

---

## Loading data from files

```python
from bags.io import load_gsd, load_cross_section, load_geometry

gsd  = load_gsd("surface_gsd.csv")            # two-column CSV: size_mm, pct_finer
xs   = load_cross_section("channel_xs.csv")   # two-column CSV: station_m, elev_m
geom = load_geometry("channel.json")          # JSON with slope, width, cross_section, …
```

## Monte Carlo / uncertainty analysis

All equation parameters are keyword arguments with scientific defaults, enabling perturbation for Monte Carlo uncertainty analysis:

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
