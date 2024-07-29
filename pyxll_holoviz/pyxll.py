"""Entry points for registering this package 
with the PyXLL Excel add-in.
"""
import pyxll
import logging
import sys

_log = logging.getLogger(__name__)


if sys.version_info[:2] >= (3, 7):
    import importlib.resources

    def _resource_bytes(package, resource_name):
        return importlib.resources.read_binary(package, resource_name)

else:
    import pkg_resources

    def _resource_bytes(package, resource_name):
        return pkg_resources.resource_stream(package, resource_name).read()


def modules():
    """Returns the modules that PyXLL should load on startup."""

    if not pyxll.__version__.endswith("dev"):
        version = tuple(map(int, pyxll.__version__.split(".")[:3]))
        if version < (5, 9, 0):
            _log.error("PyXLL version >= 5.9.0 is required to use pyxll_holoviz.")
            return []

    return [
        "pyxll_holoviz.hvplot_bridge",
        "pyxll_holoviz.panel_bridge",
        "pyxll_holoviz.ribbon",
        "pyxll_holoviz.udfs",
    ]


def ribbon():

    if not pyxll.__version__.endswith("dev"):
        version = tuple(map(int, pyxll.__version__.split(".")[:3]))
        if version < (5, 9, 0):
            _log.error("PyXLL version >= 5.9.0 is required to use pyxll_holoviz.")
            return []

    cfg = pyxll.get_config()

    disable_ribbon = False
    if cfg.has_option("HOLOVIZ", "disable_ribbon"):
        try:
            disable_ribbon = bool(int(cfg.get("HOLOVIZ", "disable_ribbon")))
        except (ValueError, TypeError):
            _log.error("Unexpected value for HOLOVIZ.disable_ribbon.")

    if disable_ribbon:
        return []

    ribbon = _resource_bytes("pyxll_holoviz.ribbon", "ribbon.xml").decode("utf-8")
    return [
        (None, ribbon)
    ]
