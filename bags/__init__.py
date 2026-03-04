"""BAGS — Bedload Analysis and Gravel Sediment Transport.

Pure-Python library implementing multiple gravel-bed bedload transport
equations from academic literature.

Quick start
-----------
>>> from bags.data import GrainSizeDistribution, ChannelGeometry
>>> from bags.equations import parker90
>>>
>>> gsd  = GrainSizeDistribution([2, 4, 8, 16, 32, 64], [5, 15, 35, 65, 85, 100])
>>> geom = ChannelGeometry(slope=0.005, width=20.0)
>>> result = parker90.transport_rate(Q=50.0, geometry=geom, surface_gsd=gsd)
>>> print(result.total_bedload_kgs, "kg/s")
"""

from .data import GrainSizeDistribution, ChannelGeometry, TransportResult
from . import equations, grain_size, hydraulics, constants, io

__version__ = "2008.11"

__all__ = [
    "GrainSizeDistribution",
    "ChannelGeometry",
    "TransportResult",
    "equations",
    "grain_size",
    "hydraulics",
    "constants",
    "io",
]
