# --- START OF FILE app.py ---

import dash
from dash import Dash, html, dcc, dash_table, callback, Output, Input, State, ctx, MATCH, ALL, ClientsideFunction
import pandas as pd
from pathlib import Path
import uuid # Import uuid for unique IDs
import math # For checking NaN

# --- Constants ---
SHEET_A_CSV = "sheetA.csv"
SHEET_B_CSV = "sheetB.csv"
MATCH_CSV = "match.csv"
TEXT_CSV = "text.csv"  
BIOGUIDE_COL = 'bioguide' # Column name for lookup key in sheet B
SEAT_COL = "seat"       # Column name in match.csv
NAME_COL = "name"       # Column name in match.csv
HIGHLIGHT_COLOR_RED = '#ffcccc'  # Light Red
HIGHLIGHT_COLOR_BLUE = '#cce5ff' # Light Blue
TEXT_TABLE_ID = 'text-table' # ID for the text data table
TEXT_FORMULA_STORE_ID = 'text-formula-store'
TEXT_SELECTION_STORE_ID = 'text-selection-mode-store'
TEXT_FORMULA_DISPLAY_ID = 'text-formula-display'
TEXT_OUTPUT_DISPLAY_ID = 'text-output-display'

# --- Helper Functions ---
def get_excel_col_name(n: int) -> str:
    """Converts a 0-based column index to an Excel-style column name (A, B, ...)."""
    name = ""
    if n < 0: return ""
    while True:
        name = chr(ord('A') + n % 26) + name
        n = n // 26 - 1
        if n < 0: break
    return name

def _to_str_safe(val):
    """Safely convert input to string, handling None and NaN."""
    if val is None: return ""
    if isinstance(val, float) and math.isnan(val): return ""
    return str(val)

def excel_left(text: str, num_chars: int) -> str:
    """Mimics Excel's LEFT function with error handling."""
    text_str = _to_str_safe(text)
    try:
        num = int(num_chars)
        if num < 0:
            return "Error: Number of characters cannot be negative."
        return text_str[:num]
    except (ValueError, TypeError):
        return "Error: Second argument (num_chars) must be a valid integer."
    except Exception as e:
        return f"Error in LEFT: {e}"

def excel_right(text: str, num_chars: int) -> str:
    """Mimics Excel's RIGHT function with error handling."""
    text_str = _to_str_safe(text)
    try:
        num = int(num_chars)
        if num < 0:
            return "Error: Number of characters cannot be negative."
        return text_str[-num:] if num > 0 else ""
    except (ValueError, TypeError):
        return "Error: Second argument (num_chars) must be a valid integer."
    except Exception as e:
        return f"Error in RIGHT: {e}"

def excel_mid(text: str, start_num: int, num_chars: int) -> str:
    """Mimics Excel's MID function with error handling."""
    text_str = _to_str_safe(text)
    try:
        start = int(start_num)
        num = int(num_chars)
        if start < 1:
            return "Error: Start number must be 1 or greater."
        if num < 0:
            return "Error: Number of characters cannot be negative."
        # Adjust start_num to be 0-based index
        return text_str[start-1 : start-1+num]
    except (ValueError, TypeError):
        return "Error: Second and third arguments must be valid integers."
    except Exception as e:
        return f"Error in MID: {e}"

def excel_substitute(text: str, old_text: str, new_text: str) -> str:
    """Mimics Excel's SUBSTITUTE function (basic version)."""
    text_str = _to_str_safe(text)
    old_text_str = _to_str_safe(old_text)
    new_text_str = _to_str_safe(new_text)
    if old_text_str == "": return text_str # Excel SUBSTITUTE returns original text if old_text is empty
    try:
        return text_str.replace(old_text_str, new_text_str)
    except Exception as e:
        return f"Error in SUBSTITUTE: {e}"

def excel_textbefore(text: str, delimiter: str) -> str:
    """Mimics Excel's TEXTBEFORE function (basic version)."""
    text_str = _to_str_safe(text)
    delimiter_str = _to_str_safe(delimiter)
    if delimiter_str == "": return "" # Excel TEXTBEFORE returns empty string if delimiter is empty
    try:
        parts = text_str.split(delimiter_str, 1)
        if len(parts) == 1:
            return f"Error: Delimiter '{delimiter_str}' not found in text."
        return parts[0]
    except Exception as e:
        return f"Error in TEXTBEFORE: {e}"

def excel_textafter(text: str, delimiter: str) -> str:
    """Mimics Excel's TEXTAFTER function (basic version)."""
    text_str = _to_str_safe(text)
    delimiter_str = _to_str_safe(delimiter)
    if delimiter_str == "": return text_str # Excel TEXTAFTER returns original text if delimiter is empty
    try:
        parts = text_str.split(delimiter_str, 1)
        if len(parts) == 1:
            return f"Error: Delimiter '{delimiter_str}' not found in text."
        return parts[1]
    except Exception as e:
        return f"Error in TEXTAFTER: {e}"

# --- Data Loading Function ---
def load_data():
    """Loads data from CSVs and preprocesses it."""
    app_dir = Path(__file__).parent
    sheet_a_path = app_dir / SHEET_A_CSV
    sheet_b_path = app_dir / SHEET_B_CSV
    match_path = app_dir / MATCH_CSV
    text_path = app_dir / TEXT_CSV 

    dataframes = {}
    errors = []

    # Load individual dataframes
    try: dataframes['a'] = pd.read_csv(sheet_a_path)
    except Exception as e: errors.append(f"Error loading {SHEET_A_CSV}: {e}")

    try: dataframes['b'] = pd.read_csv(sheet_b_path)
    except Exception as e: errors.append(f"Error loading {SHEET_B_CSV}: {e}")

    try: dataframes['match'] = pd.read_csv(match_path)
    except Exception as e: errors.append(f"Error loading {MATCH_CSV}: {e}")

    try: dataframes['text'] = pd.read_csv(text_path)
    except Exception as e: errors.append(f"Error loading {TEXT_CSV}: {e}") 

    if errors:
        # Return default empty structures on error
        print("Errors during data loading:")
        for err in errors: print(f"- {err}")
        return ({'a': pd.DataFrame(), 'b': pd.DataFrame(), 'match': pd.DataFrame(), 'text': pd.DataFrame()}, # <--- ADDED 'text' default
                {}, {}, {}, {}, -1, [], [], [], []) # <--- ADDED empty list for text cols

    df_a = dataframes['a']
    df_b = dataframes['b']
    df_match = dataframes['match']
    df_text = dataframes['text']

    # Store Original Column Lists
    original_a_cols = df_a.columns.tolist()
    original_b_cols = df_b.columns.tolist()
    original_match_cols = df_match.columns.tolist()
    original_text_cols = df_text.columns.tolist() 
    bioguide_col_index = original_b_cols.index(BIOGUIDE_COL)
    sheetB_dict_local = {row[BIOGUIDE_COL]: row.tolist() for _, row in df_b.iterrows()}
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
            bioguide_col_index, original_a_cols, original_b_cols, original_match_cols, original_text_cols) 


# --- Load Data Globally ---
try:
    # --- UNPACKING UPDATED ---
    (dfs, sheetB_dict, seatDict, nameDict, rowDict, BIOGUIDE_COL_INDEX_B,
     original_a_cols_list, original_b_cols_list, original_match_cols_list, original_text_cols_list) = load_data()
    df_a, df_b, df_match, df_text = dfs.get('a'), dfs.get('b'), dfs.get('match'), dfs.get('text') 

    # Prepare data/columns for DataTables
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

    if not df_text.empty:
        data_text = df_text.to_dict('records')
        columns_text = [{"name": i, "id": i} for i in original_text_cols_list]
    else: data_text, columns_text = [{"Error": "Load Failed"}], [{"name": "Error", "id": "Error"}]

except Exception as e:
    print(f"FATAL ERROR during data loading: {e}")
    # Set defaults for app to load without crashing
    df_a, df_b, df_match, df_text = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame() # <--- ADDED df_text default
    error_cols = [{"name": "Error", "id": "Error"}]
    error_data = [{"Error": "Data Load Failed"}]
    data_a, columns_a = error_data, error_cols
    data_b, columns_b = error_data, error_cols
    data_match, columns_match = error_data, error_cols
    data_text, columns_text = error_data, error_cols 
    sheetB_dict, seatDict, nameDict, rowDict = {}, {}, {}, {}
    BIOGUIDE_COL_INDEX_B = -1
    original_a_cols_list, original_b_cols_list, original_match_cols_list, original_text_cols_list = [], [], [], [] # <--- ADDED text default list


# --- Dash App Initialization ---
app = Dash(__name__, suppress_callback_exceptions=True, assets_folder='assets')

# --- Reusable Component Styles --- 
STYLE_DATATABLE = {'height': '250px', 'overflowY': 'auto', 'width': '100%'}
STYLE_CELL_COMMON = {'textAlign': 'left', 'padding': '5px'}
STYLE_HEADER_COMMON = {'fontWeight': 'bold'}
STYLE_CALC_BUTTON = {'marginTop': '10px'}
STYLE_RESULT_BOX = {'marginTop': '10px'}
STYLE_FORMULA_COMPONENT = {'marginRight': '5px', 'display': 'inline-block','fontFamily': 'monospace'}
STYLE_DYNAMIC_BUTTON = {'margin': '0 2px', 'fontFamily': 'monospace'}


# --- App Layout ---
app.layout = html.Div([
    html.H1("NJPC Excel Training"), # Main Title
    # Stores for text tab
    dcc.Store(id=TEXT_FORMULA_STORE_ID, data=[]), # Holds list of formula component dicts
    dcc.Store(id=TEXT_SELECTION_STORE_ID, data={'active_component_id': None, 'active_param_index': None}), # Tracks which dynamic text button is active
    dcc.Tabs(id="tab-selector", value='tab-index-match', className="tab--selector", children=[
        dcc.Tab(label='Index Match', value='tab-index-match'),
        dcc.Tab(label='Text String Basics', value='tab-text-strings'),
    ]),
    html.Div(id='tab-content')
]) # End main layout Div


# --- Callbacks ---

@app.callback(
    Output('tab-content', 'children'),
    Input('tab-selector', 'value')
)
def render_content(tab):
    if tab == 'tab-index-match':

        return html.Div([

            # === Stores for holding state ===
            dcc.Store(id='match-section-store', data={'active_button': None, 'array_col_index': None, 'array_excel_ref': None}),
            dcc.Store(id='index-section-store', data={'active_button': None, 'array_col_index': None, 'array_excel_ref': None}),
            dcc.Store(id='im-selection-mode-store', data={'active': None}),
            dcc.Store(id='im-index-param-store', data=None),
            dcc.Store(id='im-match-param-1-store', data=None),
            dcc.Store(id='im-match-param-2-store', data=None),

            # =======================================
            # === MATCH and INDEX Tutorials ===
            # =======================================
            html.Div(className="tutorial-section-container", children=[
                # --- MATCH Section ---
                html.Div(className="tutorial-section tutorial-section-match", children=[
                    html.H3("Understanding MATCH()"),
                    html.P([html.Code("MATCH(value, array, type)"), " finds the ", html.Strong("position"), " of a ", html.Strong("value"), "."]),
                    html.P("Inputs:"),
                    html.Ul([
                        html.Li([html.Code("value"), ": What you’re searching for. e.g., ", html.Code(f"{df_match.loc[0, NAME_COL] if not df_match.empty else 'Some Name'}")]),
                        html.Li([html.Code("array"), ": Which column to search. e.g., ", html.Code("B:B")]),
                        html.Li([html.Code("type"), ": Use ", html.Code("0"), " for exact match."])
                    ]),
                    html.P("Output:"),
                    html.Ul([html.Li(["The position (row number). e.g., ", html.Code("1")])]),
                    html.P(
                        "Type the value you're searching for into the 'value' box below. Then, click the 'array' button and select the column you want to search.",
                        className="instruction-text"
                    ),
                # Interactive Formula
                    html.Div(className="formula-display-interactive", children=[
                        html.Span("match(", className="formula-part-red"),
                        dcc.Input(id='match-input-value', type='text', placeholder="VALUE", size='15', className="input-box-red"),
                        html.Span(", ", className="formula-part-red"),
                        html.Button("ARRAY", id='activate-match-array', n_clicks=0, className='dynamic-text-box dynamic-text-box-red'),
                        html.Span(", 0)", className="formula-part-red")
                    ]),
                    # Table
                    dash_table.DataTable(
                        id='match-table', columns=columns_match, data=data_match,
                        column_selectable='single', selected_columns=[], cell_selectable=False, row_selectable=False, page_action='none', fixed_rows={'headers': True},
                        style_table=STYLE_DATATABLE, style_cell=STYLE_CELL_COMMON, style_header=STYLE_HEADER_COMMON
                    ),
                    # Calculate Button & Result
                    html.Button("Calculate MATCH", id='calculate-match-button', n_clicks=0, style=STYLE_CALC_BUTTON),
                    html.Div(id='match-result-display', children="Result: ", className='result-box', style=STYLE_RESULT_BOX)
                ]), # End MATCH Section Div

                # --- INDEX Section ---
                html.Div(className="tutorial-section tutorial-section-index", children=[
                    html.H3("Understanding INDEX()"),
                    html.P([html.Code("INDEX(array, position)"), " finds the ", html.Strong("value "), "at a ", html.Strong("position"), "."]),
                    html.P("Inputs:"),
                    html.Ul([
                        html.Li([html.Code("array"), ": Which column has the value you want. e.g., ", html.Code("A:A")]),
                        html.Li([html.Code("position"), ": The row number containing the value. e.g., ", html.Code("1")])
                    ]),
                    html.P("Output:"),
                    html.Ul([html.Li(["The value at that position. e.g., ", html.Code(f"{df_match.loc[0, SEAT_COL] if not df_match.empty else 'Some Seat'}")])]),
                    html.P(
                        "Click the 'array' button and select the column containing the value you want to return. Then, type the row number into the 'position' box.",
                        className="instruction-text"
                    ),
                    # Interactive Formula 
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
                        style_table=STYLE_DATATABLE, style_cell=STYLE_CELL_COMMON, style_header=STYLE_HEADER_COMMON,
                    ),
                    # Calculate Button & Result
                    html.Button("Calculate INDEX", id='calculate-index-button', n_clicks=0, style=STYLE_CALC_BUTTON),
                    html.Div(id='index-result-display', children="Result: ", className='result-box', style=STYLE_RESULT_BOX)
                ]), # End INDEX Section Div
            ]), # End Top Row Flex Container

            # =======================================
            # === INDEX/MATCH Tutorial ===
            # =======================================
            html.H2("Using INDEX() and MATCH() together"),
            html.P(["Combine ", html.Span("INDEX", style={'color':'darkblue', 'fontWeight': 'bold'}), " and ", html.Span("MATCH", style={'color':'red', 'fontWeight': 'bold'}), " to ", html.Span("look up a value from Sheet A in Sheet B", style={'color':'red'}), " and ", html.Span("return a corresponding result from the same row", style={'color':'darkblue'}), "."]),
            html.P("Instructions:", style={'fontWeight': 'bold'}),
            html.Div(className="instruction-text", children=[
                html.Ol([
                    html.Li([
                        html.Strong(html.Span("MATCH:", style={'color': 'red'})),
                        " Click the ", html.Span("'Lookup Value'", style={'color': 'red'}), " button, then select a cell in ",
                        html.Strong("Sheet A"), " containing the value you're searching for. ",
                        "Click the ", html.Span("'Lookup Column'", style={'color': 'red'}), " button, then select the column in ",
                        html.Strong("Sheet B"), " you want to search."
                    ]),
                    html.Li([
                        html.Strong(html.Span("INDEX:", style={'color': 'darkblue'})),
                        " Click the ", html.Span("'Result Column'", style={'color': 'darkblue'}), " button, then select the column in ",
                        html.Strong("Sheet B"), " containing the info you want to retrieve."
                    ])
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
                    html.H4("Sheet A", className='sheet-a-header'),
                    html.Div(className='table-container', children=[
                        dash_table.DataTable(
                            id='im-table-a', columns=columns_a, data=data_a, cell_selectable=True, fixed_rows={'headers': True},
                            row_selectable=False, column_selectable=False, page_action='none',
                            style_table=STYLE_DATATABLE, 
                            style_cell=STYLE_CELL_COMMON, style_header=STYLE_HEADER_COMMON,
                            # Conditional style added via callback
                            style_data_conditional=[]
                        )])]),
                # --- Sheet B Table ---
                html.Div(className='table-column sheet-b', children=[
                    html.H4("Sheet B", className='sheet-b-header'),
                    html.Div(className='table-container', children=[
                        dash_table.DataTable(
                            id='im-table-b', columns=columns_b, data=data_b, cell_selectable=False, fixed_rows={'headers': True},
                            row_selectable=False, column_selectable='single', selected_columns=[], page_action='none',
                            style_table=STYLE_DATATABLE,
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
        ])

    elif tab == 'tab-text-strings':
        return html.Div([
            html.H2("Text String Basics"),
            html.P("These core text functions help you extract, reshape, and combine strings in Excel. Click on a function to learn how it works and see real examples."),
            # --- Explanations ---
            html.Div(className="explanation-section", children=[
                html.Details([
                    html.Summary([html.Code("LEFT(text, num_chars)")," and ",html.Code("RIGHT(text, num_chars)")]),
                    html.P("Return a specified number of characters from the start (LEFT) or end (RIGHT) of a text string."),
                    html.Ul([
                        html.Li([html.Code("text"), ": The original text string."]),
                        html.Li([html.Code("num_chars"), ": The number of characters you want to extract."]),
                    ]),
                    html.P(["Example: ",html.Code("LEFT(\"Robert\", 3)")," returns \"Rob\" and ", html.Code("RIGHT(\"Robert\", 3)")," returns \"ert\""])
                ]),
                html.Details([
                    html.Summary(html.Code("MID(text, start_num, num_chars)")),
                    html.P("Returns a specific number of characters from a text string, starting at the position you specify."),
                    html.Ul([
                        html.Li([html.Code("text"), ": The original text string"]),
                        html.Li([html.Code("start_num"), ": The position of the first character you want to extract (1 is the first character)."]),
                        html.Li([html.Code("num_chars"), ": The number of characters you want to return."]),
                     ]),
                     html.P(["Example: ",html.Code("MID(\"Robert\", 2, 4)")," returns \"ober\""])
                ]),
                html.Details([
                    html.Summary(html.Code("SUBSTITUTE(text, old_text, new_text)")),
                    html.P("Replaces existing text with new text in a text string."),
                    html.Ul([
                        html.Li([html.Code("text"), ": The original text string."]),
                        html.Li([html.Code("old_text"), ": The text you want to replace."]),
                        html.Li([html.Code("new_text"), ": The text you want to replace OLD_TEXT with."]),
                    ]),
                    html.P(["Example: ",html.Code("SUBSTITUTE(\"Robert\", \"ert\", \"bie\")")," returns \"Robbie\""])
                 ]),
                 html.Details([
                    html.Summary([html.Code("TEXTBEFORE(text, delimiter)")," and ", html.Code("TEXTAFTER(text, delimiter)")]),
                    html.P("Return text that occurs before (TEXTBEFORE) or after (TEXTAFTER) a given character or string (delimiter)."),
                    html.Ul([
                        html.Li([html.Code("text"), ": The original text string."]),
                        html.Li([html.Code("delimeter"), ": The point before/after which you want to extract."]),
                        html.Li([html.Code("instance"), ": You can provide a third optional argument indicating which occurrence of the delimeter to use."])
                    ]),
                    html.P(["Example: ",html.Code("TEXTBEFORE(\"National Journal\", \" \")")," returns \"National\" and ",html.Code("TEXTAFTER(\"National Journal\", \" \")", " returns \"Journal\"")])
                ]),
                html.Details([
                    html.Summary([html.Code("&")]),
                    html.P("Joins several text strings into one text string."),
                    html.P(["Example: ", html.Code("\"National\" & \" \" & \"Journal\"")," returns \"National Journal\""])
                ])
            ]), # End Explanations Div

            # --- Interactive Section ---
            html.Div(className="interactive-text-section", children=[
                html.H3(["Create your own formula!"]),
                html.Div(className="instruction-text", children=[
                    html.Ul([
                        html.Li(["Click a function button to add it to the ", html.Strong("Current Formula")," below. Then, fill out its arguments to see the ",html.Strong("result"),"."]),
                        html.Li(["Use the ", html.Strong("\"\""), " button to add a text string."]),
                        html.Li(["Use the ", html.Strong("[cell]"), " button to add text directy from a cell in the table."])
                    ])
                ]),
                html.P(["Try to create names in the format: ", html.Code("Rep. Nick Begich (R-AK-AL)")]),
                
                # --- Formula Builder Buttons ---
                html.Div(className="formula-buttons", children=[
                    html.Button("LEFT", id={'type': 'add-formula-btn', 'index': 'LEFT'}, n_clicks=0),
                    html.Button("RIGHT", id={'type': 'add-formula-btn', 'index': 'RIGHT'}, n_clicks=0),
                    html.Button("MID", id={'type': 'add-formula-btn', 'index': 'MID'}, n_clicks=0),
                    html.Button("SUBSTITUTE", id={'type': 'add-formula-btn', 'index': 'SUBSTITUTE'}, n_clicks=0),
                    html.Button("TEXTBEFORE", id={'type': 'add-formula-btn', 'index': 'TEXTBEFORE'}, n_clicks=0),
                    html.Button("TEXTAFTER", id={'type': 'add-formula-btn', 'index': 'TEXTAFTER'}, n_clicks=0),
                    html.Button("&", id={'type': 'add-formula-btn', 'index': '&'}, n_clicks=0),
                    html.Button('""', id={'type': 'add-formula-btn', 'index': 'LITERAL'}, n_clicks=0),
                    html.Button('[Cell]', id={'type': 'add-formula-btn', 'index': 'CELL'}, n_clicks=0, title="Add a direct cell reference value"),
                    html.Button("Delete Last", id='delete-last-formula-btn', n_clicks=0, style={'marginLeft': '20px'}),
                    html.Button("Clear All", id='clear-formula-btn', n_clicks=0),
                ]),

                # --- Dynamic Formula Display Area ---
                html.H4("Current Formula:", style={'marginTop': '15px'}),
                html.Div(id=TEXT_FORMULA_DISPLAY_ID, className="formula-display-interactive", style={'minHeight': '40px', 'border': '1px solid #ccc', 'padding': '10px'}),

                # --- Output Display Area ---
                html.H4("Result:", style={'marginTop': '15px'}),
                html.Div(id=TEXT_OUTPUT_DISPLAY_ID, className='result-box', style={'minHeight': '30px', 'border': '1px solid #eee', 'padding': '5px', 'backgroundColor': '#f8f8f8'}),
                html.Br(),

                # --- Data Table ---
                dash_table.DataTable(
                    id=TEXT_TABLE_ID,
                    columns=columns_text,
                    data=data_text,
                    cell_selectable=True, # Allow cell selection
                    row_selectable=False,
                    column_selectable=False,
                    page_action='none',
                    fixed_rows={'headers': True},
                    style_table=STYLE_DATATABLE,
                    style_cell=STYLE_CELL_COMMON,
                    style_header=STYLE_HEADER_COMMON,
                    style_data_conditional=[], # Will be used to highlight selected cell
                    tooltip_data=[{column: {'value': str(value), 'type': 'markdown'}
                               for column, value in row.items()}
                              for row in data_text],
                     tooltip_duration=None,
                ),
            ]), # End Interactive Section Div
        ]) # End Text Tab Div

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
                'backgroundColor': HIGHLIGHT_COLOR_RED,
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
                'backgroundColor': HIGHLIGHT_COLOR_BLUE,
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

    return styles


# ==================================
# === TEXT STRING CALLBACKS      ===
# ==================================

# --- Callback to Add Formula Components ---
@callback(
    Output(TEXT_FORMULA_STORE_ID, 'data', allow_duplicate=True),
    Output(TEXT_OUTPUT_DISPLAY_ID, 'children', allow_duplicate=True), # Update output for errors
    # --- INPUTS ---
    Input({'type': 'add-formula-btn', 'index': ALL}, 'n_clicks'), # Keep this Input
    Input('clear-formula-btn', 'n_clicks'),
    Input('delete-last-formula-btn', 'n_clicks'),
    # --- STATES ---
    State(TEXT_FORMULA_STORE_ID, 'data'),
    # --- REMOVED State for add_btns_ids ---
    prevent_initial_call=True
)
# --- REMOVED add_btns_ids from signature ---
def update_formula_structure(add_btns_clicks, clear_btn_clicks, delete_btn_clicks, current_formula):
    triggered_id = ctx.triggered_id
    output_message = dash.no_update # Default to no change for error display

    # Check if the callback was triggered by any input change at all
    if not ctx.triggered:
        # print("Callback triggered but ctx.triggered is empty (e.g., initial load check).")
        return dash.no_update, dash.no_update

    triggered_input_info = ctx.triggered[0] # Info about the specific input that fired
    triggered_value = triggered_input_info['value']

    # --- Handle Clear and Delete First ---
    if triggered_id == 'clear-formula-btn':
        if triggered_value is not None and triggered_value > 0:
            print("Clear button triggered.")
            return [], "Result: Formula cleared."
        else:
            return dash.no_update, dash.no_update # Ignore if n_clicks is 0 or None

    if triggered_id == 'delete-last-formula-btn':
        if triggered_value is not None and triggered_value > 0:
            print("Delete button triggered.")
            if current_formula:
                current_formula.pop()
                return current_formula, dash.no_update # Let calculation callback update result
            else:
                return dash.no_update, "Result: Nothing to delete."
        else:
             return dash.no_update, dash.no_update # Ignore if n_clicks is 0 or None


    # --- Handle Adding Components ---
    is_add_button_click = False
    if isinstance(triggered_id, dict) and triggered_id.get('type') == 'add-formula-btn':
        if triggered_value is not None and triggered_value > 0:
            is_add_button_click = True
        # else:
        #     print(f"Ignoring add button trigger for {triggered_id}: n_clicks not > 0")


    if is_add_button_click:
        component_type = triggered_id['index']
        component_id = str(uuid.uuid4()) # Unique ID for this component instance
        print(f"Add button triggered: {component_type}")

        # --- Nesting Prevention Logic ---
        last_component_type = current_formula[-1]['type'] if current_formula else None
        can_add_value_component = not current_formula or last_component_type == 'operator'
        can_add_operator = bool(current_formula) and last_component_type != 'operator'

        new_component = None

        # --- Component Creation Logic ---
        if component_type == '&':
            if not can_add_operator:
                output_message = "Error: Cannot start with '&' or have consecutive '&&'."
            else:
                new_component = {'id': component_id, 'type': 'operator', 'value': '&'}
        elif component_type in ['LITERAL', 'CELL']:
            if not can_add_value_component:
                 output_message = f"Error: Use '&' before adding {component_type}."
            elif component_type == 'LITERAL':
                input_id = {'type': 'text-literal-input', 'index': component_id}
                new_component = {'id': component_id, 'type': 'literal_string', 'input_id': input_id, 'value': ""}
            elif component_type == 'CELL':
                button_id = {'type': 'text-cell-btn', 'index': f'{component_id}-cell'}
                new_component = {'id': component_id, 'type': 'cell_value', 'ref': None, 'value': None, 'button_id': button_id}
        elif component_type in ['LEFT', 'RIGHT', 'MID', 'SUBSTITUTE', 'TEXTBEFORE', 'TEXTAFTER']:
            if not can_add_value_component:
                output_message = f"Error: Cannot add {component_type} here. Use '&' to connect formulas or text."
            else:
                params_structure = { # Define parameters needed for each function
                    'LEFT':       [None, None], 'RIGHT':      [None, None], 'MID':        [None, None, None],
                    'SUBSTITUTE': [None, None, None], 'TEXTBEFORE': [None, None], 'TEXTAFTER':  [None, None],
                }
                param_ids_structure = { # Generate unique IDs for interactive elements
                    'LEFT':       [{'type': 'text-cell-btn', 'index': f'{component_id}-0'}, {'type': 'text-num-input', 'index': f'{component_id}-1'}],
                    'RIGHT':      [{'type': 'text-cell-btn', 'index': f'{component_id}-0'}, {'type': 'text-num-input', 'index': f'{component_id}-1'}],
                    'MID':        [{'type': 'text-cell-btn', 'index': f'{component_id}-0'}, {'type': 'text-num-input', 'index': f'{component_id}-1'}, {'type': 'text-num-input', 'index': f'{component_id}-2'}],
                    'SUBSTITUTE': [{'type': 'text-cell-btn', 'index': f'{component_id}-0'}, {'type': 'text-text-input', 'index': f'{component_id}-1'}, {'type': 'text-text-input', 'index': f'{component_id}-2'}],
                    'TEXTBEFORE': [{'type': 'text-cell-btn', 'index': f'{component_id}-0'}, {'type': 'text-text-input', 'index': f'{component_id}-1'}],
                    'TEXTAFTER':  [{'type': 'text-cell-btn', 'index': f'{component_id}-0'}, {'type': 'text-text-input', 'index': f'{component_id}-1'}],
                }
                new_component = {
                    'id': component_id, 'type': 'function', 'name': component_type,
                    'params': params_structure[component_type],
                    'param_ids': param_ids_structure[component_type]
                }

        # --- Component Handling Logic ---
        if new_component:
            current_formula.append(new_component)
            return current_formula, dash.no_update # Let calculation run
        else:
            # If component wasn't created due to error, return no_update for formula, but show error message
            return dash.no_update, output_message if output_message != dash.no_update else "Error: Invalid operation."

    # --- If not clear, delete, or a valid add button click, do nothing ---
    # print(f"No valid action taken for trigger: {triggered_id}")
    return dash.no_update, output_message

# --- Callback to Render the Dynamic Formula Display ---
@callback(
    Output(TEXT_FORMULA_DISPLAY_ID, 'children'),
    Input(TEXT_FORMULA_STORE_ID, 'data'),
    Input(TEXT_SELECTION_STORE_ID, 'data'),
)
def render_formula_display(formula_data, selection_mode):
    if not formula_data: return ""
    display_elements = []
    active_component_id = selection_mode.get('active_component_id')
    active_param_index = selection_mode.get('active_param_index')
    print(f"Rendering display. Active mode: Comp={active_component_id}, Param={active_param_index}")

    for i, component in enumerate(formula_data):
        comp_id = component['id']
        comp_type = component['type']

        # ... (operator, literal_string, cell_value rendering) ...
        if comp_type == 'operator':
            display_elements.append(html.Span(f" {component['value']} ", style=STYLE_FORMULA_COMPONENT))
        elif comp_type == 'literal_string':
            input_id = component['input_id']
            display_elements.append(html.Span('"', style=STYLE_FORMULA_COMPONENT))
            display_elements.append(dcc.Input(
                id=input_id, type='text', value=component.get('value', ''), placeholder="text",
                size='10', style=STYLE_FORMULA_COMPONENT # No debounce
            ))
            display_elements.append(html.Span('"', style=STYLE_FORMULA_COMPONENT))
        elif comp_type == 'cell_value':
             button_id = component['button_id']
             button_text = "Click to select cell"
             if isinstance(component.get('ref'), str): button_text = component['ref']
             is_active_button = (active_component_id == comp_id and active_param_index == 'cell')
             button_class = 'dynamic-text-box' + (' active' if is_active_button else '')
             print(f"  Rendering Cell Value Button: Comp={comp_id}. Mode Comp={active_component_id}, Mode Param={active_param_index}. Is Active? {is_active_button}. Class='{button_class}'")
             display_elements.append(html.Button(
                 button_text, id=button_id, n_clicks=0,
                 className=button_class, style=STYLE_DYNAMIC_BUTTON
             ))
        elif comp_type == 'function':
            func_name = component['name']
            # --- START DEBUG PRINTS ---
            # print(f"--- Processing Function Component ---")
            # print(f"DEBUG: func_name = {repr(func_name)}") 
            # print(f"DEBUG: func_name type = {type(func_name)}")
            # --- END DEBUG PRINTS ---
            params = component['params']
            param_ids = component['param_ids']
            display_elements.append(html.Span(f"{func_name}(", style=STYLE_FORMULA_COMPONENT))
            param_render_map = {
                 'LEFT':       [('cell', 'text'), ('number', '#chars')],
                 'RIGHT':      [('cell', 'text'), ('number', '#chars')],
                 'MID':        [('cell', 'text'), ('number', 'start'), ('number', '#chars')],
                 'SUBSTITUTE': [('cell', 'text'), ('text', 'old_text'), ('text', 'new_text')],
                 'TEXTBEFORE': [('cell', 'text'), ('text', 'delimiter')],
                 'TEXTAFTER':  [('cell', 'text'), ('text', 'delimiter')],
            }
            # --- START DEBUG PRINTS ---
            # print(f"DEBUG: param_render_map = {repr(param_render_map)}")
            # print(f"DEBUG: param_render_map type = {type(param_render_map)}")
            # --- END DEBUG PRINTS ---
            render_info = param_render_map.get(func_name, [])

            for p_idx, param_val in enumerate(params):
                if p_idx > 0: display_elements.append(html.Span(", ", style=STYLE_FORMULA_COMPONENT))
                param_id = param_ids[p_idx]
                param_type, placeholder = render_info[p_idx] if p_idx < len(render_info) else ('unknown', '??')

                if param_type == 'cell':
                    button_text = "Click to select cell"
                    cell_info = param_val
                    if isinstance(cell_info, dict) and 'ref' in cell_info and cell_info['ref'] is not None:
                         button_text = cell_info['ref']

                    # Check if active (compare ID and numerical index p_idx)
                    is_active_button = (active_component_id == comp_id and active_param_index == p_idx)
                    button_class = 'dynamic-text-box' + (' active' if is_active_button else '')
                    # print(f"  Button Check: Comp={comp_id}, Param={p_idx}. Mode Comp={active_component_id}, Mode Param={active_param_index}. Is Active? {is_active_button}. Class='{button_class}'")

                    display_elements.append(html.Button(
                        button_text, id=param_id, n_clicks=0,
                        className=button_class, style=STYLE_DYNAMIC_BUTTON
                    ))
                elif param_type == 'number':
                     display_elements.append(dcc.Input(
                        id=param_id, type='number', placeholder=placeholder, value=param_val,
                        min=0 if func_name in ['LEFT', 'RIGHT', 'MID'] and p_idx > 0 else None,
                        step=1, size='5',
                        style=STYLE_FORMULA_COMPONENT
                    ))
                elif param_type == 'text':
                     display_elements.append(dcc.Input(
                        id=param_id, type='text', placeholder=placeholder, value=param_val,
                        size='10',
                        style=STYLE_FORMULA_COMPONENT
                    ))

            display_elements.append(html.Span(")", style=STYLE_FORMULA_COMPONENT))

    return display_elements
# --- Callback to Activate Cell Selection Mode ---
@callback(
    Output(TEXT_SELECTION_STORE_ID, 'data'),
    Input({'type': 'text-cell-btn', 'index': ALL}, 'n_clicks'),
    State(TEXT_SELECTION_STORE_ID, 'data'),
    prevent_initial_call=True,
)
def activate_text_cell_selection(buttons_n_clicks, current_mode):
    triggered_id = ctx.triggered_id
    # Ensure a *specific* button click triggered this
    if not triggered_id or not ctx.triggered or ctx.triggered[0]['value'] is None or ctx.triggered[0]['value'] == 0:
        # print("Skipping activation: No relevant button click.")
        return dash.no_update

    print(f"Attempting activation for trigger: {triggered_id}")

    try:
        index_str = triggered_id['index']
        param_index_to_set = None
        component_id = None

        # Parse the ID to get component ID and parameter index (or 'cell')
        if index_str.endswith('-cell'):
            component_id = index_str[:-len('-cell')]
            param_index_to_set = 'cell'
        else:
            # Assume standard function parameter format: uuid-param_idx
            uuid_part, param_index_str = index_str.rsplit('-', 1)
            component_id = uuid_part
            param_index_to_set = int(param_index_str) # Can raise ValueError

        # --- If parsing succeeded, proceed ---
        print(f"Parsed OK -> Activating text selection for Component: {component_id}, Param: {param_index_to_set}")

        # Check if the mode actually needs changing
        if current_mode.get('active_component_id') != component_id or current_mode.get('active_param_index') != param_index_to_set:
             print("  -> Setting new active mode.")
             # Return the new mode state
             return {'active_component_id': component_id, 'active_param_index': param_index_to_set}
        else:
             # User clicked the *same* active button again. Don't change mode.
             print("  -> Re-clicked active button. Mode unchanged.")
             return dash.no_update

    except (ValueError, IndexError, TypeError) as e:
         # Handle errors during parsing (rsplit failure, int conversion failure)
         print(f"Error PARSING triggered_id in activate_text_cell_selection: {triggered_id}. Error: {e}")
         # Reset mode on parsing error
         return {'active_component_id': None, 'active_param_index': None}
    except Exception as e:
         # Handle any other unexpected errors during the try block
         print(f"Unexpected error in activate_text_cell_selection: {e}")
         # Reset mode on unexpected error
         return {'active_component_id': None, 'active_param_index': None}

# --- Callback to Handle Cell Selection ---
@callback(
    Output(TEXT_FORMULA_STORE_ID, 'data', allow_duplicate=True),
    Output(TEXT_SELECTION_STORE_ID, 'data', allow_duplicate=True), # Deactivate mode
    Input(TEXT_TABLE_ID, 'active_cell'),
    State(TEXT_SELECTION_STORE_ID, 'data'),
    State(TEXT_FORMULA_STORE_ID, 'data'),
    State(TEXT_TABLE_ID, 'data'), # Get current data view to map row index
    prevent_initial_call=True
)
def handle_text_cell_selection(active_cell, selection_mode, formula_data, table_data):
    active_comp_id = selection_mode.get('active_component_id')
    active_param_idx = selection_mode.get('active_param_index')

    print(f"Handle Cell Selection: active_cell={active_cell}, mode_comp={active_comp_id}, mode_param={active_param_idx}")


    if not active_cell or active_comp_id is None or active_param_idx is None:
        # print("Skipping cell update (no active cell or mode)")
        # If user clicks outside table while mode is active, deactivate mode
        if active_comp_id is not None:
             print("Deactivating cell selection mode (clicked outside).")
             return dash.no_update, {'active_component_id': None, 'active_param_index': None}
        return dash.no_update, dash.no_update


    try:
        row_index = active_cell['row']
        col_id = active_cell['column_id']

        if row_index >= len(table_data) or col_id not in original_text_cols_list:
            print(f"Error: Invalid cell coordinates: Row={row_index}, Col={col_id}")
            # Deactivate mode on error
            return dash.no_update, {'active_component_id': None, 'active_param_index': None}

        # Get value from the *currently displayed data* using row index
        cell_value = table_data[row_index].get(col_id)
        col_index = original_text_cols_list.index(col_id)
        excel_ref = f"{get_excel_col_name(col_index)}{row_index + 1}"

        cell_data = {'ref': excel_ref, 'value': cell_value}
        print(f"Selected Cell Data: {cell_data}")

        # Find the component and update the parameter
        updated = False
        for component in formula_data:
            if component['id'] == active_comp_id:
                # --- START MODIFY: Check param index type ---
                if active_param_idx == 'cell': # Handle standalone cell value component
                    if component['type'] == 'cell_value':
                         print(f"Updating cell_value component {active_comp_id} with {cell_data}")
                         component.update(cell_data) # Update ref and value directly
                         updated = True
                         break
                    else: print(f"Error: Mode indicates 'cell' but component type is {component['type']}")
                elif isinstance(active_param_idx, int): # Handle function parameter
                    if component['type'] == 'function' and 0 <= active_param_idx < len(component['params']):
                        print(f"Updating function component {active_comp_id}, param {active_param_idx} with {cell_data}")
                        component['params'][active_param_idx] = cell_data
                        updated = True
                        break
                    else: print(f"Error: Mode indicates function param but mismatch: Type={component['type']}, Index={active_param_idx}")
                else: print(f"Error: Unknown active_param_idx type: {active_param_idx}")
                # --- END MODIFY ---
                break # Found component, stop searching
        if updated:
            # Deactivate selection mode and return updated formula
            return formula_data, {'active_component_id': None, 'active_param_index': None}
        else:
            print("Error: Could not find component/param to update.")
            # Deactivate mode even if update failed
            return dash.no_update, {'active_component_id': None, 'active_param_index': None}

    except Exception as e:
        print(f"Error processing text cell selection: {e}")
        # Deactivate mode on unexpected error
        return dash.no_update, {'active_component_id': None, 'active_param_index': None}


# --- Callback to Handle Input Changes (Numbers, Text, Literals) ---
@callback(
    Output(TEXT_FORMULA_STORE_ID, 'data', allow_duplicate=True),
    # Use ALL pattern matching for dynamic inputs
    Input({'type': 'text-num-input', 'index': ALL}, 'value'),
    Input({'type': 'text-text-input', 'index': ALL}, 'value'),
    Input({'type': 'text-literal-input', 'index': ALL}, 'value'),
    State(TEXT_FORMULA_STORE_ID, 'data'),
    prevent_initial_call=True
)
def handle_text_input_change(num_values, text_values, literal_values, formula_data):
    triggered_id_dict = ctx.triggered_id
    if not triggered_id_dict or not ctx.triggered or not ctx.triggered[0]['value'] is not None: # Ensure value is not None initially
        # This check helps prevent updates on initial load where values might be None
        # print("Skipping input change: No trigger or initial None value.")
        return dash.no_update


    triggered_type = triggered_id_dict.get('type')
    triggered_index_str = triggered_id_dict.get('index') # This is comp_uuid-param_idx or comp_uuid for literal

    # Find the input value that triggered the callback
    triggered_input_prop = ctx.triggered[0]['prop_id'].split('.')[1] # 'value'
    triggered_value = ctx.triggered[0]['value']

    print(f"Input Changed: ID={triggered_id_dict}, Value={triggered_value}")


    updated = False
    try:
        if triggered_type == 'text-literal-input':
            component_id = triggered_index_str # For literals, index is just component_id
            for component in formula_data:
                if component['id'] == component_id and component['type'] == 'literal_string':
                    # Only update if value actually changed
                    if component.get('value') != triggered_value:
                         component['value'] = triggered_value
                         updated = True
                         print(f"Updated Literal Component {component_id} value to {triggered_value}")
                    break
        elif triggered_type in ['text-num-input', 'text-text-input']:
             component_id, param_index_str = triggered_index_str.rsplit('-', 1)
             param_index = int(param_index_str)

             for component in formula_data:
                 if component['id'] == component_id and component['type'] == 'function':
                     if 0 <= param_index < len(component['params']):
                         # Basic type check/conversion for numbers
                         if triggered_type == 'text-num-input':
                             try:
                                 # Allow None if input is cleared
                                 final_value = int(triggered_value) if triggered_value is not None else None
                             except (ValueError, TypeError):
                                 final_value = None # Keep as None if invalid
                                 print(f"Warning: Invalid number input '{triggered_value}' for {component_id}-{param_index}. Setting param to None.")
                         else: # text-text-input
                            final_value = triggered_value

                         # Only update if value actually changed
                         if component['params'][param_index] != final_value:
                             component['params'][param_index] = final_value
                             updated = True
                             print(f"Updated Func Component {component_id}, param {param_index} value to {final_value}")
                         break
                     else:
                          print(f"Error: Param index {param_index} out of bounds for {component_id}")
                          break # Stop searching component loop

        if updated:
            return formula_data
        else:
            # print(f"Input value for {triggered_id_dict} did not change or component not found.")
            return dash.no_update

    except (ValueError, IndexError) as e:
         print(f"Error parsing ID or index for input {triggered_id_dict}. Error: {e}")
         return dash.no_update
    except Exception as e:
        print(f"Error handling input change for {triggered_id_dict}: {e}")
        return dash.no_update

# --- Callback to Calculate and Display Final Result ---
@callback(
    Output(TEXT_OUTPUT_DISPLAY_ID, 'children', allow_duplicate=True),
    Input(TEXT_FORMULA_STORE_ID, 'data'),
    prevent_initial_call=True
)
def calculate_text_formula_result(formula_data):
    if not formula_data:
        return "Result: "

    current_result = ""
    error_occurred = False
    calculation_performed = False # Still useful to track if anything produced output

    print("\nCalculating Formula:")

    for i, component in enumerate(formula_data):
        comp_type = component['type']
        # print(f"  Processing component {i}: {comp_type}") # Keep for debugging if needed

        if error_occurred: continue # Stop calculation on first error

        if comp_type == 'operator':
            if i == 0 or formula_data[i-1]['type'] == 'operator':
                error_occurred = True
                current_result = "Error: Misplaced '&' operator."
                print(f"  Error: {current_result}")
            # If operator is last, loop ends, result is shown up to that point.
            continue

        # --- Check for missing '&' before value components ---
        if i > 0 and formula_data[i-1]['type'] != 'operator':
            error_occurred = True
            # Make error message more general
            current_result = f"Error: Missing '&' before {comp_type} component."
            print(f"  Error: {current_result}")
            continue # Stop processing if structure is wrong

        # --- Process Value Components ---
        value_to_add = None # Store the value this component contributes

        if comp_type == 'literal_string':
            value_to_add = component.get('value') # Already a string or None
            if value_to_add is None: value_to_add = "" # Treat missing value as empty string
            print(f"  Literal Value: '{value_to_add}'")

        # --- NEW: Handle 'cell_value' ---
        elif comp_type == 'cell_value':
             cell_ref_val = component.get('value')
             if cell_ref_val is None:
                 # Cell not selected yet, skip it, don't treat as error
                 print(f"  Skipping incomplete cell_value component (ID: {component['id']})")
                 continue # Move to next component
             else:
                 value_to_add = cell_ref_val # Get the stored value
                 print(f"  Cell Value: '{value_to_add}' from ref {component.get('ref')}")
        # --- END NEW ---

        elif comp_type == 'function':
            func_name = component['name']
            params = component['params']

            if any(p is None for p in params):
                # Function is incomplete, skip it, don't treat as error
                print(f"  Skipping incomplete function '{func_name}' (ID: {component['id']})")
                continue # Move to next component
            else:
                # Function is complete, try to evaluate
                processed_params = []
                param_error = False
                # ... (param processing logic) ...
                for p_idx, p_val in enumerate(params):
                    if isinstance(p_val, dict) and 'value' in p_val: processed_params.append(p_val['value'])
                    elif p_val is not None: processed_params.append(p_val)
                    else: param_error = True; break # Should not happen if initial check passed
                if param_error:
                     print(f"  Internal Error processing params for {func_name}")
                     error_occurred=True; current_result="Error: Internal Param Error."; continue


                try:
                    result_value = ""
                    # ... (call helper functions) ...
                    if func_name == 'LEFT': result_value = excel_left(*processed_params)
                    elif func_name == 'RIGHT': result_value = excel_right(*processed_params)
                    elif func_name == 'MID': result_value = excel_mid(*processed_params)
                    elif func_name == 'SUBSTITUTE': result_value = excel_substitute(*processed_params)
                    elif func_name == 'TEXTBEFORE': result_value = excel_textbefore(*processed_params)
                    elif func_name == 'TEXTAFTER': result_value = excel_textafter(*processed_params)
                    else: result_value = f"Error: Unknown function '{func_name}'"

                    print(f"  Helper func '{func_name}' returned: '{result_value}'")

                    if isinstance(result_value, str) and result_value.startswith("Error:"):
                        # Helper function returned an error - THIS IS a calculation error
                        error_occurred = True
                        current_result = result_value # Display specific error
                        print(f"  Function Helper Error: {current_result}")
                    else:
                        # Success! Store the result to be added
                         value_to_add = result_value

                except Exception as e:
                    # Error DURING calculation (e.g., wrong args passed to helper)
                    error_occurred = True
                    current_result = f"Error calculating {func_name}: {e}"
                    print(f"  Calculation Exception: {e}")

        # --- Add the result if evaluation was successful ---
        if value_to_add is not None and not error_occurred:
             current_result += _to_str_safe(value_to_add)
             calculation_performed = True
             print(f"  OK. Current Result String: '{current_result}'")


    # --- Final Output Formatting ---
    if error_occurred:
        # Display the specific error message caught during processing
        final_display = current_result
    elif not calculation_performed and not formula_data:
         final_display = "Result: " # Initial state
    elif not calculation_performed and formula_data:
         # This might happen if formula is just '&' or incomplete functions/cells
         final_display = "Result: [No output yet]"
    else:
        # Success or partial success
        final_display = f"Result: {current_result}"

    print(f"-> Final Calculation Output: {final_display}\n")
    return final_display

server = app.server

# --- Run the App ---
if __name__ == '__main__':
    # app.run(debug=True)
    app.run(debug=False) # Use False for production/deployment