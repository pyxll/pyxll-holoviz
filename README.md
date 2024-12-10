# PyXLL Holoviz Extension

Extension package for the Python Excel add-in [PyXLL](https://www.pyxll.com) to add support for [Holoviz](https://holoviz.org/) charts and panels.

This package enables displaying HoloViews, hvPlot, and panel objects in Excel via the PyXLL function
[pyxll.plot](https://www.pyxll.com/docs/userguide/plotting/index.html).

## Requirements

- PyXLL >= 5.9.0

## Installation

To install, run the following in the same Python environment that the PyXLL add-in is configured to use:

```
pip install pyxll-holoviz
```

For installing the Excel add-in PyXLL see the [PyXLL user guide](https://www.pyxll.com/docs/userguide/installation/firsttime.html).

## Examples

The following is a simple example taken from the [Holoviz hvplot examples](https://hvplot.holoviz.org/).

It uses ``hvplot`` to create a scatter plot figure, and then uses [pyxll.plot](https://www.pyxll.com/docs/userguide/plotting/index.html) to display that figure in Excel using PyXLL.

```python
import pyxll
from bokeh.sampledata.penguins import data as df
import hvplot.pandas

@pyxll.xl_func
def test_plot():
    # Create the holoviews figure using hvplot
    figure = df.hvplot.scatter(x='bill_length_mm',
                               y='bill_depth_mm',
                               by='species',
                               responsive=True)

    # Show the figure in Excel
    pyxll.plot(figure)
```

When this function is called from Excel as a worksheet function, the figure is displayed as an interactive chart in Excel. The kwarg ``reponsive=True`` is used so that the figure scales automatically if resized in Excel.

This next example shows how to display a panel in Excel and is taken from the [panel basic tutorial](https://panel.holoviz.org/tutorials/basic/serve.html).

```python
import pyxll
import panel as pn

@pyxll.xl_func
def test_panel():
    # Create a holoviz panel object
    panel = pn.panel("Hello World")

    # Show the panel in Excel
    pyxll.plot(panel)
```

When this function is called from Excel as a worksheet function, the panel is displayed as an interactive widget in Excel below the cell where the function was called.

## Extras

As well as supporting displaying Holoviz objects via [pyxll.plot](https://www.pyxll.com/docs/userguide/plotting/index.html) this package also adds a right click context menu ``hvPlot Explorer`` and a new ``hvplot`` UDF (worksheet function).

### hvPlot Explorer

To use the hvPlot explorer, select a tabular range of data in Excel including the column headers. Then right click to bring up the right click context menu and select ``hvPlot Explorer``.

This will add the Holoviz explorer panel to the worksheet as an interactive control.

### hvplot UDF

The worksheet function ``hvplot`` plots a pandas DataFrame using hvplot.

Parameters are passed in as a dictionary of keyword arguemnts (kwargs), and additionally
as pairs of (key, value).

kwargs can be passed as the second argument as an array, or as additional arguments with the key as one argument followed by the value as the next argument. This is is to allow for passing values are are arrays that
can be passed as part of the first kwargs array.

For example, to call the equivalent of the Python code:

```python
df.hvplot(kind='scatter', x='index', y=['value1', 'value2'])
```

The Excel formula would be the following, where the input DataFrame
is in cell A1:

```
=hvplot(A1,
        {"kind", "scatter"; "x", "index"},
        "y", {"value1", "value2"})
```

## Config

The custom context menu (ribbon) component of this extension can be diabled by adding the following
to your pyxll.cfg config file:

```
[HOLOVIZ]
disable_ribbon = 1
```
