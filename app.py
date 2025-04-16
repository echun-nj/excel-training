import dash
from dash import Dash, html, dcc, dash_table, callback, Output, Input, State, ctx
import pandas as pd
from pathlib import Path

# --- Constants ---
SHEET_A_CSV = "sheetA.csv"
SHEET_B_CSV = "sheetB.csv"
MATCH_CSV = "match.csv"
BIOGUIDE_COL = 'bioguide' # Column name for lookup key in sheet B
SEAT_COL = "seat"       # Column name in match.csv
NAME_COL = "name"       # Column name in match.csv
HIGHLIGHT_COLOR_RED = '#ffcccc'  # Light Red
HIGHLIGHT_COLOR_BLUE = '#cce5ff' # Light Blue

# --- Helper Function ---
def get_excel_col_name(n: int) -> str:
    """Converts a 0-based column index to an Excel-style column name (A, B, ...)."""
    name = ""
    if n < 0: return ""
    while True:
        name = chr(ord('A') + n % 26) + name
        n = n // 26 - 1
        if n < 0: break
    return name

# --- Data Loading Function ---
def load_data():
    """Loads data from CSVs and preprocesses it."""
    app_dir = Path(__file__).parent
    sheet_a_path = app_dir / SHEET_A_CSV
    sheet_b_path = app_dir / SHEET_B_CSV
    match_path = app_dir / MATCH_CSV

    dataframes = {}
    errors = []

    # Load individual dataframes
    try: dataframes['a'] = pd.read_csv(sheet_a_path)
    except Exception as e: errors.append(f"Error loading {SHEET_A_CSV}: {e}")

    try: dataframes['b'] = pd.read_csv(sheet_b_path)
    except Exception as e: errors.append(f"Error loading {SHEET_B_CSV}: {e}")

    try: dataframes['match'] = pd.read_csv(match_path)
    except Exception as e: errors.append(f"Error loading {MATCH_CSV}: {e}")

    if errors:
        # Return default empty structures on error
        print("Errors during data loading:")
        for err in errors: print(f"- {err}")
        return ({'a': pd.DataFrame(), 'b': pd.DataFrame(), 'match': pd.DataFrame()},
                {}, {}, {}, {}, -1, [], [], [])

    df_a = dataframes['a']
    df_b = dataframes['b']
    df_match = dataframes['match']

    # Store Original Column Lists
    original_a_cols = df_a.columns.tolist()
    original_b_cols = df_b.columns.tolist()
    original_match_cols = df_match.columns.tolist()
    print(f"Original A Columns: {original_a_cols}")
    print(f"Original B Columns: {original_b_cols}")
    print(f"Original Match Columns: {original_match_cols}")

    # Create sheetB dictionary
    if BIOGUIDE_COL not in original_b_cols:
        raise ValueError(f"'{BIOGUIDE_COL}' column not found in {SHEET_B_CSV}.")
    bioguide_col_index = original_b_cols.index(BIOGUIDE_COL)
    print(f"Bioguide Index (0-based): {bioguide_col_index}")
    # Use original df_b for dictionary values
    sheetB_dict_local = {row[BIOGUIDE_COL]: row.tolist() for _, row in df_b.iterrows()}

    # Create match.csv dictionaries
    if SEAT_COL not in original_match_cols:
        raise ValueError(f"'{SEAT_COL}' column not found in {MATCH_CSV}.")
    if NAME_COL not in original_match_cols:
        raise ValueError(f"'{NAME_COL}' column not found in {MATCH_CSV}.")

    seatDict_local = {}
    nameDict_local = {}
    rowDict_local = {}
    for index, row in df_match.iterrows():
        row_num = index + 1
        seat_val = row[SEAT_COL]
        name_val = row[NAME_COL]
        seatDict_local[seat_val] = row_num
        nameDict_local[name_val] = row_num
        rowDict_local[row_num] = [seat_val, name_val]

    return (dataframes, sheetB_dict_local, seatDict_local, nameDict_local, rowDict_local,
            bioguide_col_index, original_a_cols, original_b_cols, original_match_cols)

# --- Load Data Globally ---
try:
    (dfs, sheetB_dict, seatDict, nameDict, rowDict, BIOGUIDE_COL_INDEX_B,
     original_a_cols_list, original_b_cols_list, original_match_cols_list) = load_data()
    df_a, df_b, df_match = dfs.get('a'), dfs.get('b'), dfs.get('match')

    # Prepare data/columns for DataTables (can be done once)
    if not df_a.empty:
        data_a = df_a.to_dict('records')
        columns_a = [{"name": i, "id": i} for i in original_a_cols_list]
    else: data_a, columns_a = [{"Error": "Load Failed"}], [{"name": "Error", "id": "Error"}]

    if not df_b.empty:
        data_b = df_b.to_dict('records')
        columns_b = [{"name": i, "id": i, "selectable": True} for i in original_b_cols_list]
    else: data_b, columns_b = [{"Error": "Load Failed"}], [{"name": "Error", "id": "Error"}]

    if not df_match.empty:
        data_match = df_match.to_dict('records')
        columns_match = [{"name": i, "id": i, "selectable": True} for i in original_match_cols_list]
    else: data_match, columns_match = [{"Error": "Load Failed"}], [{"name": "Error", "id": "Error"}]

except Exception as e:
    print(f"FATAL ERROR during data loading: {e}")
    # Set defaults for app to load without crashing
    df_a, df_b, df_match = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    error_cols = [{"name": "Error", "id": "Error"}]
    error_data = [{"Error": "Data Load Failed"}]
    data_a, columns_a = error_data, error_cols
    data_b, columns_b = error_data, error_cols
    data_match, columns_match = error_data, error_cols
    sheetB_dict, seatDict, nameDict, rowDict = {}, {}, {}, {}
    BIOGUIDE_COL_INDEX_B = -1
    original_a_cols_list, original_b_cols_list, original_match_cols_list = [], [], []


# --- Dash App Initialization ---
app = Dash(__name__, suppress_callback_exceptions=True, assets_folder='assets')

# --- Reusable Component Styles --- (Used directly in layout)
# DataTable styles (keeping background colors here for reliability)
STYLE_DATATABLE = {'height': '200px', 'overflowY': 'auto', 'width': '100%'}
STYLE_DATATABLE_INDEXMATCH_A = {'height': '400px', 'overflowY': 'auto', 'width': '100%', 'backgroundColor': '#ffebeb'}
STYLE_DATATABLE_INDEXMATCH_B = {'height': '400px', 'overflowY': 'auto', 'width': '100%', 'backgroundColor': '#e0f2f7'}
STYLE_DATATABLE_TUTORIAL = {'height': '200px', 'overflowY': 'auto', 'width': '100%', 'backgroundColor': '#e0f2f7'} # Shared style for tutorial tables
STYLE_CELL_COMMON = {'textAlign': 'left', 'padding': '5px'}
STYLE_HEADER_COMMON = {'fontWeight': 'bold'}
STYLE_CALC_BUTTON = {'marginTop': '10px'}
STYLE_RESULT_BOX = {'marginTop': '10px'}


# --- App Layout ---
app.layout = html.Div([
    # === Stores for holding state ===
    dcc.Store(id='match-section-store', data={'active_button': None, 'array_col_index': None, 'array_excel_ref': None}),
    dcc.Store(id='index-section-store', data={'active_button': None, 'array_col_index': None, 'array_excel_ref': None}),
    dcc.Store(id='im-selection-mode-store', data={'active': None}), # Renamed for clarity (im = index/match)
    dcc.Store(id='im-index-param-store', data=None),
    dcc.Store(id='im-match-param-1-store', data=None),
    dcc.Store(id='im-match-param-2-store', data=None),

    html.H1("NJPC Excel Training"), # Main Title

    # =======================================
    # === MATCH and INDEX Tutorials ===
    # =======================================
    html.Div(className="tutorial-section-container", children=[
        # --- MATCH Section ---
        html.Div(className="tutorial-section tutorial-section-match", children=[
            html.H3("Understanding MATCH()"),
            html.P([html.Code("MATCH(VALUE, ARRAY, TYPE)"), " finds the ", html.Strong("position"), " of a ", html.Strong("value"), "."]),
            html.P("Inputs:"),
            html.Ul([
                html.Li([html.Strong("VALUE:"), " What you’re searching for. e.g., ", html.Code(f"{df_match.loc[0, NAME_COL] if not df_match.empty else 'Some Name'}")]),
                html.Li([html.Strong("ARRAY:"), " Which column to search. e.g., ", html.Code("B:B")]),
                html.Li([html.Strong("TYPE:"), " Use ", html.Code("0"), " for exact match."])
            ]),
            html.P("Output:"),
            html.Ul([html.Li(["The position (row number). e.g., ", html.Code("1")])]),
            html.P(
                "Type the value you're searching for into the 'VALUE' box below. Then, click the 'ARRAY' button and select the column you want to search.",
                className="instruction-text"
            ),
        # Interactive Formula
            html.Div(className="formula-display-interactive", children=[
                html.Span("MATCH(", className="formula-part-red"),
                dcc.Input(id='match-input-value', type='text', placeholder="VALUE", size='15', className="input-box-red"),
                html.Span(", ", className="formula-part-red"),
                html.Button("ARRAY", id='activate-match-array', n_clicks=0, className='dynamic-text-box dynamic-text-box-red'),
                html.Span(", 0)", className="formula-part-red")
            ]),
            # Table
            dash_table.DataTable(
                id='match-table', columns=columns_match, data=data_match,
                column_selectable='single', selected_columns=[], cell_selectable=False, row_selectable=False, page_action='none', fixed_rows={'headers': True},
                style_table=STYLE_DATATABLE_TUTORIAL, style_cell=STYLE_CELL_COMMON, style_header=STYLE_HEADER_COMMON
            ),
            # Calculate Button & Result
            html.Button("Calculate MATCH", id='calculate-match-button', n_clicks=0, style=STYLE_CALC_BUTTON),
            html.Div(id='match-result-display', children="Result: ", className='result-box', style=STYLE_RESULT_BOX)
        ]), # End MATCH Section Div

        # --- INDEX Section ---
        html.Div(className="tutorial-section tutorial-section-index", children=[
            html.H3("Understanding INDEX()"),
            html.P([html.Code("INDEX(ARRAY, POSITION)"), " finds the ", html.Strong("value "), "at a ", html.Strong("position"), "."]),
            html.P("Inputs:"),
            html.Ul([
                html.Li([html.Strong("ARRAY:"), " Which column has the value you want. e.g., ", html.Code("A:A")]),
                html.Li([html.Strong("POSITION:"), " The row number containing the value. e.g., ", html.Code("1")])
            ]),
            html.P("Output:"),
            html.Ul([html.Li(["The value at that position. e.g., ", html.Code(f"{df_match.loc[0, SEAT_COL] if not df_match.empty else 'Some Seat'}")])]),
            html.P(
                "Click the 'ARRAY' button and select the column containing the value you want to return. Then, type the row number into the 'POSITION' box.",
                className="instruction-text"
            ),
            # Interactive Formula - APPLY BLUE STYLES
            html.Div(className="formula-display-interactive", children=[
                # Use className for color
                html.Span("INDEX(", className="formula-part-blue"),
                # Use className for blue border
                html.Button("ARRAY", id='activate-index-array', n_clicks=0, className='dynamic-text-box dynamic-text-box-blue'),
                html.Span(", ", className="formula-part-blue"),
                # Use className for blue border
                dcc.Input(id='index-input-position', type='number', placeholder="POSITION", min=1, step=1, size='10', className="input-box-blue"),
                html.Span(")", className="formula-part-blue")
            ]),
             # Table
            dash_table.DataTable(
                id='index-table', columns=columns_match, data=data_match,
                column_selectable='single', selected_columns=[], cell_selectable=False, row_selectable=False, page_action='none', fixed_rows={'headers': True},
                style_table=STYLE_DATATABLE_TUTORIAL, style_cell=STYLE_CELL_COMMON, style_header=STYLE_HEADER_COMMON,
            ),
             # Calculate Button & Result
            html.Button("Calculate INDEX", id='calculate-index-button', n_clicks=0, style=STYLE_CALC_BUTTON),
            html.Div(id='index-result-display', children="Result: ", className='result-box', style=STYLE_RESULT_BOX)
        ]), # End INDEX Section Div
    ]), # End Top Row Flex Container

    # --- Separator ---
    html.Hr(className="section-separator"),

    # =======================================
    # === INDEX/MATCH Tutorial ===
    # =======================================
    html.H2("Using INDEX() and MATCH() together"),
    html.P(["Combine ", html.Span("INDEX", style={'color':'blue'}), " and ", html.Span("MATCH", style={'color':'red'}), " to ", html.Span("look up a value from Sheet A in Sheet B", style={'color':'red'}), " and ", html.Span("return a corresponding result from the same row", style={'color':'blue'}), "."]),
    html.P("Instructions:", style={'fontWeight': 'bold'}),
    html.Div(className="instruction-text", children=[
        html.P([
            "1. ", 
            html.Strong(html.Span("MATCH:", style={'color': 'red'})), # Label is red and bold
            " Click the ", html.Span("'Lookup Value'", style={'color':'red'}), " button, then select a cell in ", html.Strong("Sheet A"), " containing the value you're searching for. ",
            "Click the ", html.Span("'Lookup Column'", style={'color':'red'}), " button, then select the column  in ", html.Strong("Sheet B"), " you want to search."
        ]),
        html.P([
            "2. ",
            html.Strong(html.Span("INDEX:", style={'color': 'darkblue'})), # Label is blue and bold
            " Click the ", html.Span("'Result Column'", style={'color':'darkblue'}), " button, then select the column in ", html.Strong("Sheet B"), " containing the info you want to retrieve."
        ])
    ]),
    # --- Formula Display ---
    html.Div(className='formula-display', children=[
        html.Span("INDEX(", className="formula-part-blue"),
        html.Span("sheetB!", className="formula-part-blue"),
        # Button 1: Blue
        html.Button("Result Column", id='im-activate-dyn1', n_clicks=0, className='dynamic-text-box dynamic-text-box-blue'),
        html.Span(", ", className="formula-part-blue"),
        html.Span("MATCH(", className="formula-part-red"),
        # Button 2: Red
        html.Button("Lookup Value", id='im-activate-dyn2', n_clicks=0, className='dynamic-text-box dynamic-text-box-red'),
        html.Span(", ", className="formula-part-red"),
        html.Span("sheetB!", className="formula-part-red"),
         # Button 3: Red
        html.Button("Lookup Column", id='im-activate-dyn3', n_clicks=0, className='dynamic-text-box dynamic-text-box-red'),
        html.Span(", 0)", className="formula-part-red"),
        html.Span(")", className="formula-part-blue")
    ]),

    # --- Tables Side-by-Side ---
    html.Div(className="index-match-tables-container", children=[
        # --- Sheet A Table ---
        html.Div(className='table-column sheet-a', children=[
            html.H4("Sheet A", className='sheet-a-header'), # Use class for black text
            html.Div(className='table-container', children=[
                dash_table.DataTable(
                    id='im-table-a', columns=columns_a, data=data_a, cell_selectable=True, fixed_rows={'headers': True},
                    row_selectable=False, column_selectable=False, page_action='none',
                    style_table=STYLE_DATATABLE_INDEXMATCH_A, # Red background
                    style_cell=STYLE_CELL_COMMON, style_header=STYLE_HEADER_COMMON,
                    # Conditional style added via callback
                    style_data_conditional=[]
                )])]),
        # --- Sheet B Table ---
        html.Div(className='table-column sheet-b', children=[
             html.H4("Sheet B", className='sheet-b-header'), # Use class for black text
             html.Div(className='table-container', children=[
                 dash_table.DataTable(
                    id='im-table-b', columns=columns_b, data=data_b, cell_selectable=False, fixed_rows={'headers': True},
                    row_selectable=False, column_selectable='single', selected_columns=[], page_action='none',
                    style_table=STYLE_DATATABLE_INDEXMATCH_B, # Blue background
                    style_cell={**STYLE_CELL_COMMON, 'minWidth': '100px'}, style_header=STYLE_HEADER_COMMON,
                     # Conditional style added via callback
                    style_data_conditional=[]
                 )])])]),

    # --- Calculate Button ---
    html.Div(children=[ 
        html.Button("Calculate INDEX/MATCH Result", id='im-calculate-button', n_clicks=0)
    ]),

    # --- Result Display ---
    html.Div(className="index-match-result-container", children=[
        html.Div(id='im-result-display', children="Result: ", className='result-box')
    ]),

    html.P([
        " Once you've built an INDEX/MATCH formula in Excel for one row, like this, you can drag the formula down and dynamically perform the same lookup for all other rows!"
    ]) 
]) # End main layout Div


# --- Callbacks ---

# ==========================
# === MATCH CALLBACKS ======
# ==========================
@callback(
    Output('match-section-store', 'data', allow_duplicate=True), # Use allow_duplicate
    Input('activate-match-array', 'n_clicks'),
    State('match-section-store', 'data'),
    prevent_initial_call=True
)
def activate_match_array_selection(n_clicks, current_store_data):
    """Activates selection mode for MATCH array button."""
    if n_clicks > 0:
        print("MATCH activate button clicked")
        current_store_data['active_button'] = 'activate-match-array'
        return current_store_data
    return dash.no_update

@callback(
    Output('activate-match-array', 'className'),
    Input('match-section-store', 'data')
)
def style_match_array_button(match_store_data):
    """Updates style of MATCH array button based on active state."""
    base_class = "dynamic-text-box dynamic-text-box-red"
    active = match_store_data and match_store_data.get('active_button') == 'activate-match-array'
    return f"{base_class}{' active' if active else ''}"

@callback(
    Output('activate-match-array', 'children'),
    Output('match-section-store', 'data', allow_duplicate=True),
    Input('match-table', 'selected_columns'),
    State('match-section-store', 'data'),
    prevent_initial_call=True
)
def handle_match_column_selection(selected_columns, current_store_data):
    """Handles column selection in MATCH table, updates store and button."""
    active_button = current_store_data.get('active_button')
    button_id_to_match = 'activate-match-array'
    print(f"MATCH Table Selected Columns: {selected_columns}, Current Mode: {active_button}")

    if active_button != button_id_to_match or not selected_columns or not original_match_cols_list:
        print("Skipping MATCH column update")
        return dash.no_update, dash.no_update

    try:
        selected_col_id = selected_columns[0]
        if selected_col_id not in original_match_cols_list:
             print(f"Error: Column '{selected_col_id}' not found in original Match cols.")
             current_store_data['active_button'] = None # Reset mode on error
             return dash.no_update, current_store_data

        col_index = original_match_cols_list.index(selected_col_id)
        if col_index not in [0, 1]: # Specific check for match table columns
            print(f"Error: Invalid column index ({col_index}) selected from match table.")
            current_store_data['active_button'] = None
            return dash.no_update, current_store_data

        # Update store
        col_letter = get_excel_col_name(col_index)
        excel_col_ref = f"{col_letter}:{col_letter}"
        current_store_data['array_col_index'] = col_index
        current_store_data['array_excel_ref'] = excel_col_ref
        current_store_data['active_button'] = None # Deactivate
        print(f"MATCH array selected: Col={selected_col_id}, Idx={col_index}, Ref={excel_col_ref}")
        return excel_col_ref, current_store_data # Return button text, updated store

    except Exception as e: # Catch broader exceptions
         print(f"Error processing MATCH column selection: {e}")
         current_store_data['active_button'] = None # Reset mode on error
         return dash.no_update, current_store_data

@callback(
    Output('match-result-display', 'children'),
    Input('calculate-match-button', 'n_clicks'),
    State('match-input-value', 'value'),
    State('match-section-store', 'data'),
    prevent_initial_call=True
)
def calculate_match_result(n_clicks, lookup_value, match_store_data):
    """Calculates and displays the MATCH result."""
    selected_col_idx = match_store_data.get('array_col_index')
    print(f"Calculating MATCH: Value='{lookup_value}', ColIdx={selected_col_idx}")

    result_val = ""
    if selected_col_idx is None: result_val = "Error: Select ARRAY column."
    elif not lookup_value: result_val = "Error: Enter VALUE."

    elif selected_col_idx == original_match_cols_list.index(SEAT_COL):
        result_val = seatDict.get(lookup_value, "#N/A (no match found)")
    elif selected_col_idx == original_match_cols_list.index(NAME_COL):
         result_val = nameDict.get(lookup_value, "#N/A (no match found)")
    else: result_val = "Error: Invalid column selected."

    return f"Result: {result_val}"

@callback(
    Output('match-table', 'style_data_conditional'),
    Input('match-section-store', 'data') # Trigger based on the store's data
)

def style_selected_match_column(match_store_data):
    """Applies highlight style based on the column index stored for MATCH."""
    styles = []
    if not match_store_data: return styles
    col_index = match_store_data.get('array_col_index')
    if col_index is not None and original_match_cols_list and 0 <= col_index < len(original_match_cols_list):
        try:
            selected_id = original_match_cols_list[col_index]
            styles.append({
                'if': {'column_id': selected_id},
                'backgroundColor': HIGHLIGHT_COLOR_RED, # CHANGE TO RED HIGHLIGHT
                'color': 'black'
            })
        except Exception as e: print(f"Error styling MATCH col: {e}")
    return styles

# ==========================
# === INDEX CALLBACKS ======
# ==========================
@callback(
    Output('index-section-store', 'data', allow_duplicate=True),
    Input('activate-index-array', 'n_clicks'),
    State('index-section-store', 'data'),
    prevent_initial_call=True
)
def activate_index_array_selection(n_clicks, current_store_data):
    """Activates selection mode for INDEX array button."""
    if n_clicks > 0:
        print("INDEX activate button clicked")
        current_store_data['active_button'] = 'activate-index-array'
        return current_store_data
    return dash.no_update

@callback(
    Output('activate-index-array', 'className'),
    Input('index-section-store', 'data')
)
def style_index_array_button(index_store_data):
    """Updates style of INDEX array button based on active state."""
    base_class = "dynamic-text-box dynamic-text-box-blue"
    active = index_store_data and index_store_data.get('active_button') == 'activate-index-array'
    return f"{base_class}{' active' if active else ''}"

@callback(
    Output('activate-index-array', 'children'),
    Output('index-section-store', 'data', allow_duplicate=True),
    Input('index-table', 'selected_columns'),
    State('index-section-store', 'data'),
    prevent_initial_call=True
)
def handle_index_column_selection(selected_columns, current_store_data):
    """Handles column selection in INDEX table, updates store and button."""
    active_button = current_store_data.get('active_button')
    button_id_to_match = 'activate-index-array'
    print(f"INDEX Table Selected Columns: {selected_columns}, Current Mode: {active_button}")

    if active_button != button_id_to_match or not selected_columns or not original_match_cols_list:
        print("Skipping INDEX column update")
        return dash.no_update, dash.no_update

    try:
        selected_col_id = selected_columns[0]
        if selected_col_id not in original_match_cols_list:
             print(f"Error: Column '{selected_col_id}' not found in original Match cols.")
             current_store_data['active_button'] = None
             return dash.no_update, current_store_data

        col_index = original_match_cols_list.index(selected_col_id)
        if col_index not in [0, 1]:
            print(f"Error: Invalid column index ({col_index}) selected from index table.")
            current_store_data['active_button'] = None
            return dash.no_update, current_store_data

        # Update store
        col_letter = get_excel_col_name(col_index)
        excel_col_ref = f"{col_letter}:{col_letter}"
        current_store_data['array_col_index'] = col_index
        current_store_data['array_excel_ref'] = excel_col_ref
        current_store_data['active_button'] = None # Deactivate
        print(f"INDEX array selected: Col={selected_col_id}, Idx={col_index}, Ref={excel_col_ref}")
        return excel_col_ref, current_store_data # Return button text, updated store

    except Exception as e:
         print(f"Error processing INDEX column selection: {e}")
         current_store_data['active_button'] = None
         return dash.no_update, current_store_data

@callback(
    Output('index-result-display', 'children'),
    Input('calculate-index-button', 'n_clicks'),
    State('index-input-position', 'value'),
    State('index-section-store', 'data'),
    prevent_initial_call=True
)
def calculate_index_result(n_clicks, position_input, index_store_data):
    """Calculates and displays the INDEX result."""
    selected_col_idx = index_store_data.get('array_col_index')
    print(f"Calculating INDEX: Position='{position_input}', ColIdx={selected_col_idx}")

    result_val = ""
    if selected_col_idx is None: result_val = "Error: Select ARRAY column."
    elif position_input is None: result_val = "Error: Enter POSITION number."
    else:
        try:
            position = int(position_input)
            if position <= 0: raise ValueError("Position must be positive.")

            if position in rowDict:
                row_data_list = rowDict[position] # List [seat, name]
                # Use constants for column indices
                seat_idx = original_match_cols_list.index(SEAT_COL)
                name_idx = original_match_cols_list.index(NAME_COL)

                if selected_col_idx == seat_idx: result_val = row_data_list[seat_idx] # Should be 0
                elif selected_col_idx == name_idx: result_val = row_data_list[name_idx] # Should be 1
                else: result_val = "Error: Invalid internal column index." # Should not happen
            else:
                result_val = f"#N/A (position {position} not found)"

        except ValueError: result_val = "Error: Position must be a positive number."
        except Exception as e:
            print(f"Unexpected Error during INDEX calculation: {e}")
            result_val = f"Error: {e}"

    return f"Result: {result_val}"

@callback(
    Output('index-table', 'style_data_conditional'),
    Input('index-section-store', 'data') # Trigger based on the INDEX store's data
)

def style_selected_index_column(index_store_data):
    """Applies highlight style based on the column index stored for INDEX."""
    styles = []
    if not index_store_data: return styles
    col_index = index_store_data.get('array_col_index')
    if col_index is not None and original_match_cols_list and 0 <= col_index < len(original_match_cols_list):
        try:
            selected_id = original_match_cols_list[col_index]
            styles.append({
                'if': {'column_id': selected_id},
                'backgroundColor': HIGHLIGHT_COLOR_BLUE, # CHANGE TO BLUE HIGHLIGHT
                'color': 'black'
            })
        except Exception as e: print(f"Error styling INDEX col: {e}")
    return styles


# ==================================
# === INDEX/MATCH CALLBACKS ===
# ==================================
@callback(
    Output('im-selection-mode-store', 'data'),
    Input('im-activate-dyn1', 'n_clicks'),
    Input('im-activate-dyn2', 'n_clicks'),
    Input('im-activate-dyn3', 'n_clicks'),
    prevent_initial_call=True
)
def update_indexmatch_selection_mode(n1, n2, n3):
    """Activates selection mode for INDEX/MATCH buttons."""
    button_id = ctx.triggered_id
    # Map button ID to mode number
    mode_map = {'im-activate-dyn1': 1, 'im-activate-dyn2': 2, 'im-activate-dyn3': 3}
    new_mode = mode_map.get(button_id)
    if new_mode:
        print(f"INDEX/MATCH Activation: Mode {new_mode}")
        return {'active': new_mode}
    return dash.no_update

@callback(
    Output('im-activate-dyn1', 'className'), Output('im-activate-dyn2', 'className'), Output('im-activate-dyn3', 'className'),
    Input('im-selection-mode-store', 'data'))

def update_indexmatch_button_styles(store):
    """Updates styles for INDEX/MATCH activation buttons using red/blue scheme."""
    mode = store.get('active') if store else None
    # Assign classes based on function: INDEX=blue, MATCH=red
    cls = {
        1: "dynamic-text-box dynamic-text-box-blue", # Dyn1 (INDEX Array) = Blue
        2: "dynamic-text-box dynamic-text-box-red",  # Dyn2 (MATCH Value) = Red
        3: "dynamic-text-box dynamic-text-box-red"   # Dyn3 (MATCH Array) = Red
    }
    # Apply 'active' class if mode matches
    return (f"{cls[1]}{' active' if mode == 1 else ''}",
            f"{cls[2]}{' active' if mode == 2 else ''}",
            f"{cls[3]}{' active' if mode == 3 else ''}")


@callback(
    Output('im-activate-dyn2', 'children'),    # Button for MATCH value
    Output('im-match-param-1-store', 'data'), # Store for value + ref
    Output('im-selection-mode-store', 'data', allow_duplicate=True), # Reset mode
    Input('im-table-a', 'active_cell'),
    State('im-selection-mode-store', 'data'),
    prevent_initial_call=True
)
def handle_im_sheet_a_cell_selection(active_cell, selection_mode_data):
    """Handles cell selection in Sheet A for INDEX/MATCH."""
    mode = selection_mode_data.get('active') if selection_mode_data else None
    print(f"IM Table A Active Cell: {active_cell}, Current Mode: {mode}")

    if mode != 2 or not active_cell or not original_a_cols_list or df_a.empty:
        print("Skipping IM sheet A update")
        return dash.no_update, dash.no_update, dash.no_update

    try:
        row_index = active_cell['row'] # 0-based index in the displayed data
        col_id = active_cell['column_id']

        if col_id not in original_a_cols_list:
             print(f"Error: IM Column '{col_id}' not found in original Sheet A.")
             return "Error: Col?", None, {'active': None} # Reset mode
        col_index = original_a_cols_list.index(col_id)

        # Calculate Excel ref using 1-based row index
        col_letter = get_excel_col_name(col_index)
        excel_ref = f"{col_letter}{row_index + 1}"

        # Get value from dataframe using iloc
        if row_index >= df_a.shape[0] or col_index >= df_a.shape[1]:
             print(f"Error: IM Invalid index for df_a. Row: {row_index}, Col: {col_index}")
             return "Error: Idx?", None, {'active': None} # Reset mode
        cell_value = df_a.iloc[row_index, col_index]

        # Store data and update button
        match_param_data = {'cell_ref': excel_ref, 'cell_value': cell_value}
        print(f"IM Sheet A selected: Ref={excel_ref}, Val={cell_value}, ColIdx={col_index}")
        return excel_ref, match_param_data, {'active': None} # Reset mode

    except Exception as e:
        print(f"Error processing IM Sheet A selection: {e}")
        return "Error", None, {'active': None} # Reset mode


@callback(
    Output('im-activate-dyn1', 'children'),    # Button for INDEX array
    Output('im-activate-dyn3', 'children'),    # Button for MATCH array
    Output('im-index-param-store', 'data'),    # Store for INDEX col
    Output('im-match-param-2-store', 'data'),  # Store for MATCH col
    Output('im-selection-mode-store', 'data', allow_duplicate=True), # Reset mode
    Input('im-table-b', 'selected_columns'),
    State('im-selection-mode-store', 'data'),
    prevent_initial_call=True
)
def handle_im_sheet_b_column_selection(selected_columns, selection_mode_data):
    """Handles column selection in Sheet B for INDEX/MATCH."""
    mode = selection_mode_data.get('active') if selection_mode_data else None
    print(f"IM Table B Selected Columns: {selected_columns}, Current Mode: {mode}")

    if mode not in [1, 3] or not selected_columns or not original_b_cols_list:
        print("Skipping IM sheet B update")
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update

    try:
        selected_col_id = selected_columns[0]
        if selected_col_id not in original_b_cols_list:
             print(f"Error: IM Column '{selected_col_id}' not found in original Sheet B.")
             return dash.no_update, dash.no_update, dash.no_update, dash.no_update, {'active': None}

        col_index = original_b_cols_list.index(selected_col_id)

        # Prepare common data
        col_letter = get_excel_col_name(col_index)
        excel_col_ref = f"{col_letter}:{col_letter}"
        param_data = {'col_index': col_index, 'excel_ref': excel_col_ref}
        print(f"IM Sheet B selected: Col={selected_col_id}, Idx={col_index}, Ref={excel_col_ref}, Mode={mode}")

        # Initialize outputs to no_update
        out_dyn1, out_dyn3 = dash.no_update, dash.no_update
        out_idx_param, out_match_param2 = dash.no_update, dash.no_update

        # Update the correct store and button based on mode
        if mode == 1:
            out_dyn1 = excel_col_ref
            out_idx_param = param_data
        elif mode == 3:
            out_dyn3 = excel_col_ref
            out_match_param2 = param_data

        return out_dyn1, out_dyn3, out_idx_param, out_match_param2, {'active': None} # Reset mode

    except Exception as e:
         print(f"Error processing IM Sheet B selection: {e}")
         return dash.no_update, dash.no_update, dash.no_update, dash.no_update, {'active': None}


@callback(
    Output('im-result-display', 'children'),
    Input('im-calculate-button', 'n_clicks'),
    State('im-index-param-store', 'data'),
    State('im-match-param-1-store', 'data'),
    State('im-match-param-2-store', 'data'),
    prevent_initial_call=True
)
def calculate_im_result(n_clicks, index_data, match1_data, match2_data):
    """Calculates and displays the final INDEX/MATCH result."""
    print(f"Calculating INDEX/MATCH: Index={index_data}, Match1={match1_data}, Match2={match2_data}")

    result_val = "" # Use a single variable for the final output string

    if not index_data or not match1_data or not match2_data:
        result_val = "Error: Please select all parts of the formula."
    else:
        try:
            idx_param = index_data['col_index']
            m_param1_val = match1_data['cell_value']
            m_param2_idx = match2_data['col_index']
            print(f"  Params: IdxCol={idx_param}, MatchVal='{m_param1_val}', MatchCol={m_param2_idx}, ExpectBioIdx={BIOGUIDE_COL_INDEX_B}")

            if BIOGUIDE_COL_INDEX_B == -1:
                result_val = "Config Error: Bioguide index not found."
            elif m_param2_idx != BIOGUIDE_COL_INDEX_B:
                bio_col = original_b_cols_list[BIOGUIDE_COL_INDEX_B] if original_b_cols_list and 0 <= BIOGUIDE_COL_INDEX_B < len(original_b_cols_list) else BIOGUIDE_COL
                sel_col = original_b_cols_list[m_param2_idx] if original_b_cols_list and 0 <= m_param2_idx < len(original_b_cols_list) else f'Col {m_param2_idx}'
                bio_ref = match2_data.get('excel_ref', f"{get_excel_col_name(BIOGUIDE_COL_INDEX_B)}:{get_excel_col_name(BIOGUIDE_COL_INDEX_B)}")
                sel_ref = match2_data.get('excel_ref', f"{get_excel_col_name(m_param2_idx)}:{get_excel_col_name(m_param2_idx)}")
                result_val = f"Error: Your lookup column does not contain the lookup value. Try choosing another column."
            else:
                matched_row_list = sheetB_dict.get(m_param1_val)
                if matched_row_list is None:
                    bio_col = original_b_cols_list[BIOGUIDE_COL_INDEX_B] if original_b_cols_list else BIOGUIDE_COL
                    result_val = f"No match found for '{m_param1_val}' in '{bio_col}' column."
                elif 0 <= idx_param < len(matched_row_list):
                    final_value = matched_row_list[idx_param]
                    result_val = "[Blank]" if pd.isna(final_value) else str(final_value)
                else:
                     result_val = f"Error: Result column index ({idx_param}) out of bounds (max {len(matched_row_list)-1})."

        except KeyError as e: result_val = f"Calc Error: Missing data '{e}'. Select all parts."
        except Exception as e:
            print(f"Unexpected Error during INDEX/MATCH calculation: {e}")
            result_val = f"Unexpected Error: {e}"

    return f"Result: {result_val}"

@callback(
    Output('im-table-b', 'style_data_conditional'),
    Input('im-index-param-store', 'data'),  # INDEX col index (for BLUE)
    Input('im-match-param-2-store', 'data') # MATCH col index (for RED)
)
def style_selected_im_b_columns(index_param_data, match_param_2_data):
    """Applies BLUE highlight to INDEX col, RED highlight to MATCH col in Sheet B."""
    styles = []
    index_col_idx = index_param_data.get('col_index') if index_param_data else None
    match_col_idx = match_param_2_data.get('col_index') if match_param_2_data else None

    print(f"Styling IM B: IndexCol={index_col_idx}, MatchCol={match_col_idx}")

    # Helper to add style if index is valid
    def add_style(col_idx, color):
        if col_idx is not None and original_b_cols_list and 0 <= col_idx < len(original_b_cols_list):
            try:
                col_id = original_b_cols_list[col_idx]
                print(f"  Applying {'RED' if color == HIGHLIGHT_COLOR_RED else 'BLUE'} to '{col_id}'")
                styles.append({
                    'if': {'column_id': col_id},
                    'backgroundColor': color,
                    'color': 'black'
                })
            except Exception as e: print(f"Error adding style: {e}")

    # Apply BLUE for INDEX column FIRST
    add_style(index_col_idx, HIGHLIGHT_COLOR_BLUE)

    # Apply RED for MATCH column SECOND (will override blue if same column)
    add_style(match_col_idx, HIGHLIGHT_COLOR_RED)

    print(f"-> Final B Styles: {styles}")
    return styles

server = app.server

# --- Run the App ---
if __name__ == '__main__':
    # app.run(debug=True)
    app.run(debug=False)