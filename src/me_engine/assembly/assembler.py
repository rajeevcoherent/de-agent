"""The deterministic assembler — the math spine.

Given a `DriverSet`, it computes the Value, ASP and Volume bands for every
geography by applying the core identities (verified against the source file):

    segment_value = band_total * segment_share          (split)
    volume        = value / asp * 1000                  (price -> volume)
    market_share  = segment / band_total                (view)

There is no spreadsheet code here and no per-cell looping in client code — all
year-wise arithmetic is delegated to `Series`, so each rule reads as one line.
"""
from __future__ import annotations

from typing import Mapping

from ..domain.drivers import DriverSet, GeographyDrivers
from ..domain.series import Series
from ..domain.runtime_taxonomy import RuntimeTaxonomy, default_taxonomy
from ..domain.taxonomy import Band, Dimension
from .model import BandResult, GeographyResult, MarketResult, MetricRow

VOLUME_UNIT_SCALE = 1000.0     # US$ Mn / (US$ per unit) -> Th Liters


class Assembler:
    """Turns drivers into a fully-resolved `MarketResult`."""

    def __init__(self, taxonomy: RuntimeTaxonomy | None = None) -> None:
        self._taxonomy = taxonomy or default_taxonomy()

    def assemble(self, drivers: DriverSet) -> MarketResult:
        results = {
            name: self._assemble_geography(geo)
            for name, geo in drivers.geographies.items()
        }
        return MarketResult(market_name=drivers.market_name, geographies=results)

    # --- per geography ------------------------------------------------------
    def _assemble_geography(self, geo: GeographyDrivers) -> GeographyResult:
        value_band = self._value_band(geo)
        asp_band = self._asp_band(geo)
        volume_band = self._volume_band(value_band, asp_band, geo)
        bands = {Band.VALUE: value_band, Band.ASP: asp_band, Band.VOLUME: volume_band}
        return GeographyResult(name=geo.name, bands=bands)

    def _value_band(self, geo: GeographyDrivers) -> BandResult:
        rows = self._segment_rows(geo.value, geo)
        return BandResult(band=Band.VALUE, total=geo.value, rows_by_label=rows)

    def _asp_band(self, geo: GeographyDrivers) -> BandResult:
        """ASP only spans the leaf products of the priced dimension.

        Parent/rollup segments are excluded — they have no meaningful ASP row
        in the workbook (their cell is either 0 or a sum). Only leaf products
        participate in ASP and Volume calculations.
        """
        priced = self._taxonomy.priced_dimension
        parents_in_priced = {
            priced.parent_of(s) for s in priced.segments if priced.parent_of(s) is not None
        }
        leaf_products = [p for p in priced.segments if p not in parents_in_priced]

        rows = {}
        for product in leaf_products:
            try:
                series = geo.asp.for_product(product)
            except KeyError:
                continue
            rows[product] = MetricRow(product, series, share_of_parent=None)

        if not rows:
            # Fallback: use whatever the driver has
            for product in priced.segments:
                try:
                    rows[product] = MetricRow(product, geo.asp.for_product(product),
                                              share_of_parent=None)
                except KeyError:
                    pass

        placeholder_total = next(iter(rows.values())).series if rows else geo.value
        return BandResult(band=Band.ASP, total=placeholder_total, rows_by_label=rows)

    def _volume_band(
        self, value_band: BandResult, asp_band: BandResult, geo: GeographyDrivers,
    ) -> BandResult:
        """Volume per product = value / asp * 1000; band total = sum of products.

        Only products that appear in BOTH the value band and ASP band are included.
        """
        common_products = [
            p for p in asp_band.rows_by_label
            if p in value_band.rows_by_label
        ]
        if not common_products:
            # Degenerate case — return empty volume band
            return BandResult(band=Band.VOLUME, total=geo.value, rows_by_label={})

        product_rows = {
            product: MetricRow(
                product,
                value_band.rows_by_label[product].series.divided_by(
                    asp_band.rows_by_label[product].series, VOLUME_UNIT_SCALE),
                share_of_parent=None,
            )
            for product in common_products
        }
        total_volume = self._sum_series(r.series for r in product_rows.values())
        rows = self._attach_shares(product_rows, total_volume)
        return BandResult(band=Band.VOLUME, total=total_volume, rows_by_label=rows)

    # --- segmentation expansion --------------------------------------------
    def _segment_rows(self, total: Series, geo: GeographyDrivers) -> dict[str, MetricRow]:
        """Expand every segmentation dimension into absolute value rows.

        A segment's *value* is its share-of-immediate-parent times that parent's
        value (the band total for top-level segments, or a sibling segment for
        hierarchical channels). The stored `share_of_parent` is exactly the
        driver share — i.e. relative to the immediate parent — matching how the
        source workbook renders the Market Share columns.
        """
        rows: dict[str, MetricRow] = {}
        for dim in self._taxonomy.segmentation_dimensions:
            shares = geo.segmentation.for_dimension(dim)
            for segment in dim.segments:
                parent_label = dim.parent_of(segment)
                parent_value = total if parent_label is None else rows[parent_label].series
                series = parent_value.scaled_by(shares[segment])
                rows[segment] = MetricRow(segment, series, share_of_parent=shares[segment])
        return rows

    @staticmethod
    def _attach_shares(rows: Mapping[str, MetricRow], total: Series) -> dict[str, MetricRow]:
        return {
            label: MetricRow(label, row.series, share_of_parent=row.series.share_of(total))
            for label, row in rows.items()
        }

    @staticmethod
    def _sum_series(series_iter) -> Series:
        series_list = list(series_iter)
        first = series_list[0]
        return Series({
            year: sum(s.values[year] for s in series_list)
            for year in first.values
        })
