"""
Code injected (silently) into every kernel right after it starts.

It does two things:

  1. Registers a custom display formatter so that *any* time a pandas
     DataFrame / Series is the last expression in a cell, the kernel also
     emits a compact JSON representation under a private MIME type. The
     task pane recognises that MIME type and offers to push it to the grid.

  2. Injects an ``xl`` helper into the user namespace so a cell can
     explicitly target a location:

         xl.write(df, anchor="B2", sheet="Report", table=True)

Both paths funnel through the same serializer, so behaviour is identical.
"""

DF_MIME = "application/vnd.jupyxl.dataframe+json"

BOOTSTRAP = r'''
import json as _json

_JUPYXL_DF_MIME = "application/vnd.jupyxl.dataframe+json"
_JUPYXL_MAX_ROWS = 100_000          # hard cap so we never ship a billion rows over a socket
_JUPYXL_MAX_COLS = 4_000


def _jupyxl_payload(obj):
    """Turn a DataFrame (or Series) into a JSON-safe dict the task pane understands."""
    import pandas as pd

    if isinstance(obj, pd.Series):
        df = obj.to_frame()
    else:
        df = obj

    truncated_rows = len(df) > _JUPYXL_MAX_ROWS
    truncated_cols = df.shape[1] > _JUPYXL_MAX_COLS
    view = df.iloc[:_JUPYXL_MAX_ROWS, :_JUPYXL_MAX_COLS]

    columns = [str(c) for c in view.columns]
    dtypes = [str(t) for t in view.dtypes]

    # to_json handles NaN -> null, datetimes -> ISO, numpy ints/floats -> JSON numbers.
    rows = _json.loads(view.to_json(orient="values", date_format="iso", default_handler=str))

    return {
        "columns": columns,
        "dtypes": dtypes,
        "rows": rows,
        "nrows": int(len(df)),
        "ncols": int(df.shape[1]),
        "truncated_rows": truncated_rows,
        "truncated_cols": truncated_cols,
    }


def _jupyxl_register_formatter():
    try:
        ip = get_ipython()
    except NameError:
        return
    if ip is None:
        return

    from IPython.core.formatters import BaseFormatter

    formatters = ip.display_formatter.formatters
    if _JUPYXL_DF_MIME not in formatters:
        class _JupyXLFormatter(BaseFormatter):
            format_type = _JUPYXL_DF_MIME
        formatters[_JUPYXL_DF_MIME] = _JupyXLFormatter()

    fmt = formatters[_JUPYXL_DF_MIME]
    # Registering by name means pandas does not have to be imported yet.
    fmt.for_type_by_name("pandas.core.frame", "DataFrame", _jupyxl_payload)
    fmt.for_type_by_name("pandas.core.series", "Series", _jupyxl_payload)


class _JupyXL:
    """Explicit grid-targeting helper available to user code as ``xl``."""

    def write(self, df, anchor="A1", sheet=None, table=True):
        """Send a DataFrame to Excel at a chosen anchor.

        Parameters
        ----------
        df      : pandas DataFrame or Series
        anchor  : top-left cell, e.g. "A1" or "C5"
        sheet   : worksheet name; None -> active sheet (created if missing)
        table   : wrap the output in a styled Excel Table
        """
        from IPython.display import display

        payload = _jupyxl_payload(df)
        payload["target"] = {"anchor": str(anchor), "sheet": sheet, "table": bool(table)}
        display({_JUPYXL_DF_MIME: payload}, raw=True)
        return None

    def __repr__(self):
        return "<jupyxl bridge: xl.write(df, anchor='A1', sheet=None, table=True)>"


try:
    _jupyxl_register_formatter()
except Exception as _e:  # pandas missing is fine — xl.write will still error clearly
    pass

xl = _JupyXL()
'''
