"""Optional CSV / JSON loaders for channel geometry and grain-size data.

The core library never touches files; these helpers are thin convenience
wrappers for loading :class:`~bags.data.GrainSizeDistribution` and
:class:`~bags.data.ChannelGeometry` objects from disk.
"""
from __future__ import annotations
import csv
import json
import re
from pathlib import Path

from .data import GrainSizeDistribution, ChannelGeometry


def _sniff_delimiter(text: str) -> str:
    """Return the most likely column delimiter for a CSV snippet."""
    for delim in (",", "\t", ";", " "):
        if delim in text:
            return delim
    return ","


def load_gsd(path: str | Path) -> GrainSizeDistribution:
    """Load a grain-size distribution from a two-column CSV file.

    The file may have an optional header row.  Columns are:
    ``grain_size_mm``, ``percent_finer``.

    Auto-detects the delimiter (comma, tab, semicolon, or space).

    Parameters
    ----------
    path : str or Path

    Returns
    -------
    GrainSizeDistribution
    """
    text = Path(path).read_text()
    delim = _sniff_delimiter(text.split("\n")[0])
    reader = csv.reader(text.splitlines(), delimiter=delim)

    sizes, finer = [], []
    for row in reader:
        row = [c.strip() for c in row if c.strip()]
        if len(row) < 2:
            continue
        try:
            s = float(row[0])
            f = float(row[1])
        except ValueError:
            continue          # header row
        sizes.append(s)
        finer.append(f)

    return GrainSizeDistribution(sizes_mm=sizes, finer_pct=finer)


def load_cross_section(path: str | Path) -> list[tuple[float, float]]:
    """Load a channel cross-section from a two-column CSV file.

    Columns: ``station_m``, ``elevation_m``.
    Auto-detects delimiter.

    Returns
    -------
    list[tuple[float, float]]
    """
    text = Path(path).read_text()
    delim = _sniff_delimiter(text.split("\n")[0])
    reader = csv.reader(text.splitlines(), delimiter=delim)

    pts: list[tuple[float, float]] = []
    for row in reader:
        row = [c.strip() for c in row if c.strip()]
        if len(row) < 2:
            continue
        try:
            x = float(row[0])
            z = float(row[1])
        except ValueError:
            continue
        pts.append((x, z))

    return pts


def load_geometry(path: str | Path) -> ChannelGeometry:
    """Load a :class:`~bags.data.ChannelGeometry` from a JSON file.

    Expected JSON keys (all optional except ``slope``):

    .. code-block:: json

        {
          "slope": 0.005,
          "width": 20.0,
          "mannings_n": 0.035,
          "mannings_n_left": null,
          "mannings_n_right": null,
          "cross_section": [[0, 1.5], [5, 0.0], [20, 1.5]]
        }

    ``cross_section`` is a list of ``[station, elevation]`` pairs.
    """
    data = json.loads(Path(path).read_text())

    xs_raw = data.get("cross_section")
    xs: list[tuple[float, float]] | None = None
    if xs_raw is not None:
        xs = [(float(pt[0]), float(pt[1])) for pt in xs_raw]

    return ChannelGeometry(
        slope=float(data["slope"]),
        width=float(data["width"]) if data.get("width") is not None else None,
        cross_section=xs,
        mannings_n=float(data["mannings_n"]) if data.get("mannings_n") else None,
        mannings_n_left=float(data["mannings_n_left"]) if data.get("mannings_n_left") else None,
        mannings_n_right=float(data["mannings_n_right"]) if data.get("mannings_n_right") else None,
    )
