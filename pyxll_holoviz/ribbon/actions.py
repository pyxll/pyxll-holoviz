from pyxll import XLCell, ObjectCacheKeyError, xl_app, schedule_call, xlcAlert, plot
from functools import wraps


def alert_on_error(func):
    """Decorator to show an error using xlcAlert on error.

    Can only be used for functions called from a macro or
    using pyxll.schedule_call.
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            xlcAlert(f"An error occurred: {e}\n\n"
                     "Check the PyXLL log file for more details.")
            raise

    return wrapper


def hvplot_explorer(control):
    """Ribbon action for the 'hvPlot Explorer' button.

    Get the DataFrame from the current selection and opens
    an hvPlot Explorer panel.
    """
    schedule_call(_open_hvplot_explorer)


@alert_on_error
def _open_hvplot_explorer():
    """
    Get the current selection and display an hvPlot explorer.
    Must be called from a macro, or using pyxll.schedule_call.
    """
    # Avoid importing in the module as it slows down loading Excel
    import pandas as pd

    # Get the current selection
    xl = xl_app(com_package="win32com")
    cell = XLCell.from_range(xl.Selection)

    # Check if the value is a cached object
    try:
        df = cell.options(type="object").value
        if not isinstance(df, pd.DataFrame):
            raise ValueError(f"Expected a pandas DataFrame, got '{df.__class__.__name__}'.")
    except ObjectCacheKeyError:
        df = None
        pass

    # If we didn't find a DataFrame as a cached object maybe it's a range of data
    if df is None:
        df = cell.options(type="pandas.dataframe<columns=True, index=False>", auto_resize=True).value

    # Make sure the hvplot pandas extension is loaded
    import hvplot.pandas  # noqa

    # Get the explorer panel
    explorer = df.hvplot.explorer(responsive=True)

    # Get the Range from the cell (which may be expanded) and get the position of the
    # next cell to the right.
    r = cell.to_range(com_package="win32com")
    top_right = r.Item(1, r.Columns.Count+1)

    # And show it in Excel as an embedded plot control
    plot(explorer,
         top=top_right.Top,
         left=top_right.Left,
         width=800,
         height=500)
