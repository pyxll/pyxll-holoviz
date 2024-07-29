"""
Bridge class that allows PyXLL to display holoviz plots
by saving them as html.
"""
from pyxll import PlotBridgeBase, xl_plot_bridge


@xl_plot_bridge(
    "holoviews.*",
    "hvplot.*"
)
class HvPlotBridge(PlotBridgeBase):

    def __init__(self, figure):
        PlotBridgeBase.__init__(self, figure)

    def can_export(self, format):
        return format == "html"

    def get_size_hint(self, dpi):
        opts = self.figure.opts
        width = getattr(opts, "width", None)
        height = getattr(opts, "height", None)
        if width is None or height is None:
            return None
        return width * 72.0 / dpi, height * 72.0 / dpi

    def export(self, width, height, dpi, format, filename, **kwargs):
        if format != "html":
            raise ValueError("Unable to export as '%s'" % str(format))

        # Import holoviews here to avoid slowing down Excel when loading
        import holoviews
        holoviews.save(self.figure, filename)
