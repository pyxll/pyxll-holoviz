"""
Worksheet functions for displaying plots with hvPlot.
"""
from pyxll import xl_func, plot, get_type_converter


@xl_func("pandas.dataframe df, optional<dict<str, var>> kwargs, var *kv_pairs: object",
         name="hvplot",
         category="Holoviz")
def hvplot(df, kwargs, *kv_pairs):
    """
    Plot a pandas DataFrame using hvplot.

    Parameters are passed in as a dictionary of keyword arguemnts, and additionally
    as pairs of key, value. For example, to call the equivalent of the Python code::

        df.hvplot(kind='scatter', x='index', y=['value1', 'value2'])

    The Excel formula would be the following, where the input DataFrame
    is in cell A1::

        =hvplot(A1, {"kind", "scatter"; "x", "index"}, "y", {"value1", "value2"})
    """
    if kwargs is None:
        kwargs = {}

    if kv_pairs:
        if len(kv_pairs) % 2 != 0:
            raise ValueError("Mismatched key value pairs (should be an even number)")
        
        for key, value in zip(kv_pairs[::2], kv_pairs[1::2]):
            # Skip any missing keys
            if key is None:
                continue

            # Convert any 2d ranges into flattened lists
            if isinstance(value, list):
                to_list = get_type_converter("var", "var[]")
                value = to_list(value)

            kwargs[str(key)] = value

    # Make sure the hvplot pandas extension is loaded
    import hvplot.pandas  # noqa

    # Create the chart and display it in Excel
    kwargs.setdefault("responsive", True)
    chart = df.hvplot(**kwargs)
    plot(chart)

    return chart
