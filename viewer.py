"""
Audit Report Viewer
===================
Simple UI to browse and open downloaded audit reports.
Run: python viewer.py  → open http://localhost:8052
"""

import os
from datetime import datetime
import dash
from dash import dcc, html, dash_table, Input, Output, State
import pandas as pd

DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

app = dash.Dash(__name__, title="Audit Report Viewer")

app.layout = html.Div(style={"fontFamily": "Arial, sans-serif", "maxWidth": "900px",
                              "margin": "30px auto", "padding": "0 20px"}, children=[

    html.H2("QA Suite Audit Reports", style={"color": "#2c3e50"}),

    html.Div(style={"display": "flex", "gap": "10px", "marginBottom": "16px"}, children=[
        html.Button("Refresh", id="refresh-btn", n_clicks=0,
                    style={"padding": "8px 20px", "backgroundColor": "#2c3e50",
                           "color": "white", "border": "none", "borderRadius": "4px",
                           "cursor": "pointer", "fontSize": "14px"}),
        html.Button("Open Selected File", id="open-btn", n_clicks=0,
                    style={"padding": "8px 20px", "backgroundColor": "#27ae60",
                           "color": "white", "border": "none", "borderRadius": "4px",
                           "cursor": "pointer", "fontSize": "14px"}),
        html.Button("Open Downloads Folder", id="folder-btn", n_clicks=0,
                    style={"padding": "8px 20px", "backgroundColor": "#7f8c8d",
                           "color": "white", "border": "none", "borderRadius": "4px",
                           "cursor": "pointer", "fontSize": "14px"}),
    ]),

    html.Div(id="status-msg",
             style={"marginBottom": "12px", "color": "#27ae60", "fontSize": "14px"}),

    dash_table.DataTable(
        id="file-table",
        columns=[
            {"name": "File Name",    "id": "name"},
            {"name": "Size",         "id": "size"},
            {"name": "Date Downloaded", "id": "modified"},
        ],
        data=[],
        row_selectable="single",
        selected_rows=[],
        style_table={"overflowX": "auto"},
        style_cell={"textAlign": "left", "padding": "10px 12px",
                    "fontFamily": "Arial", "fontSize": "13px"},
        style_header={"backgroundColor": "#2c3e50", "color": "white",
                      "fontWeight": "bold", "fontSize": "13px"},
        style_data_conditional=[
            {"if": {"row_index": "odd"}, "backgroundColor": "#f8f9fa"},
            {"if": {"state": "selected"}, "backgroundColor": "#d6eaf8",
             "border": "1px solid #2980b9"},
        ],
        style_as_list_view=True,
    ),

    dcc.Store(id="file-paths-store"),  # hidden store for full file paths
])


def get_file_list():
    rows, paths = [], []
    if os.path.isdir(DOWNLOAD_DIR):
        files = [
            (f, os.path.join(DOWNLOAD_DIR, f))
            for f in os.listdir(DOWNLOAD_DIR)
            if os.path.isfile(os.path.join(DOWNLOAD_DIR, f))
        ]
        files.sort(key=lambda x: os.path.getmtime(x[1]), reverse=True)
        for fname, fpath in files:
            stat = os.stat(fpath)
            size_kb = stat.st_size / 1024
            size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.2f} MB"
            rows.append({
                "name":     fname,
                "size":     size_str,
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d  %H:%M"),
            })
            paths.append(fpath)
    return rows, paths


@app.callback(
    Output("file-table", "data"),
    Output("file-paths-store", "data"),
    Output("status-msg", "children"),
    Input("refresh-btn", "n_clicks"),
)
def refresh(n):
    rows, paths = get_file_list()
    if not rows:
        return [], [], f"No files found in {DOWNLOAD_DIR}"
    return rows, paths, f"{len(rows)} file(s) in {DOWNLOAD_DIR}"


@app.callback(
    Output("status-msg", "children", allow_duplicate=True),
    Input("open-btn", "n_clicks"),
    State("file-table", "selected_rows"),
    State("file-paths-store", "data"),
    prevent_initial_call=True,
)
def open_file(n, selected_rows, paths):
    if not selected_rows:
        return "Please select a file first."
    if not paths:
        return "No files loaded — click Refresh."
    idx = selected_rows[0]
    path = paths[idx]
    if os.path.isfile(path):
        os.startfile(path)
        return f"Opened: {os.path.basename(path)}"
    return f"File not found: {path}"


@app.callback(
    Output("status-msg", "children", allow_duplicate=True),
    Input("folder-btn", "n_clicks"),
    prevent_initial_call=True,
)
def open_folder(n):
    import subprocess
    subprocess.Popen(f'explorer "{DOWNLOAD_DIR}"')
    return f"Opened folder: {DOWNLOAD_DIR}"


if __name__ == "__main__":
    rows, _ = get_file_list()
    print(f"Viewing: {DOWNLOAD_DIR}  ({len(rows)} file(s))")
    print("Open http://localhost:8052 in your browser")
    app.run(debug=False, port=8052)
