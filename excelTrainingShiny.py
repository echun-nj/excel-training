from shiny import App, ui, reactive, render

# Sample (filler) data for the basic example
basic_data = [
    {"A": "California", "B": "Sunny"},
    {"A": "New York",   "B": "Busy"},
    {"A": "Texas",      "B": "Hot"},
    {"A": "Florida",    "B": "Warm"},
    {"A": "Illinois",   "B": "Windy"}
]

# Sample data for the complex example:
# Sheet 1 contains the lookup value (cell A2) and Sheet 2 contains the lookup and index arrays.
sheet1_data = [
    {"A": "California"}
]

sheet2_data = [
    {"C": "California", "B": "Los Angeles"},
    {"C": "New York",   "B": "New York City"},
    {"C": "Texas",      "B": "Houston"},
    {"C": "Florida",    "B": "Miami"},
    {"C": "Illinois",   "B": "Chicago"}
]

app_ui = ui.page_fluid(
    ui.tags.h2("Interactive Excel Training: Index Match"),
    # Include JavaScript to capture cell clicks.
    ui.tags.script("""
      // This function is called when a cell is clicked.
      // It sends an input value (named by the 'example' prefix) with the cell info.
      function cellClicked(example, row, col, value) {
          Shiny.setInputValue(example + '_selected', {row: row, col: col, value: value}, {priority: "event"});
      }
    """),
    
    # ==============================
    # Basic Example Section
    # ==============================
    ui.h3("Basic Example: INDEX(B:B, MATCH(\"California\", A:A, 0))"),
    # A div to display the working formula (it will update reactively)
    ui.div(ui.HTML("<div id='basic_formula'>Formula will appear here</div>")),
    ui.br(),
    # The basic example table (filler data)
    ui.HTML("""
      <table id="basic_table" border="1" style="border-collapse: collapse;">
        <thead>
          <tr><th>A</th><th>B</th></tr>
        </thead>
        <tbody>
          <tr>
            <td onclick="cellClicked('basic', 1, 'A', 'California')">California</td>
            <td>Sunny</td>
          </tr>
          <tr>
            <td onclick="cellClicked('basic', 2, 'A', 'New York')">New York</td>
            <td>Busy</td>
          </tr>
          <tr>
            <td onclick="cellClicked('basic', 3, 'A', 'Texas')">Texas</td>
            <td>Hot</td>
          </tr>
          <tr>
            <td onclick="cellClicked('basic', 4, 'A', 'Florida')">Florida</td>
            <td>Warm</td>
          </tr>
          <tr>
            <td onclick="cellClicked('basic', 5, 'A', 'Illinois')">Illinois</td>
            <td>Windy</td>
          </tr>
        </tbody>
      </table>
    """),
    ui.br(),
    # Updated button using ui.input_action_button
    ui.input_action_button("basic_show", "Show Basic Result"),
    ui.div(ui.HTML("<div id='basic_result'>Result will appear here</div>")),
    
    ui.hr(),
    
    # ==============================
    # Complex Example Section
    # ==============================
    ui.h3("Complex Example: INDEX([Sheet 2]!B:B, MATCH(A2, [Sheet 2]!C:C, 0))"),
    # Div to display the complex formula (reactively updated)
    ui.div(ui.HTML("<div id='complex_formula'>Formula will appear here</div>")),
    ui.br(),
    ui.h4("Sheet 1 (Select the Lookup Value for A2)"),
    # Sheet 1 table: only one cell is provided here for demonstration.
    ui.HTML("""
      <table id="sheet1_table" border="1" style="border-collapse: collapse;">
        <thead>
          <tr><th>A</th></tr>
        </thead>
        <tbody>
          <tr>
            <td onclick="cellClicked('sheet1', 2, 'A', 'California')">California</td>
          </tr>
        </tbody>
      </table>
    """),
    ui.br(),
    ui.h4("Sheet 2 (Data Table)"),
    # Sheet 2 table: shows two columns, where column C is used for MATCH and column B for INDEX.
    ui.HTML("""
      <table id="sheet2_table" border="1" style="border-collapse: collapse;">
        <thead>
          <tr><th>C</th><th>B</th></tr>
        </thead>
        <tbody>
          <tr>
            <td>California</td>
            <td>Los Angeles</td>
          </tr>
          <tr>
            <td>New York</td>
            <td>New York City</td>
          </tr>
          <tr>
            <td>Texas</td>
            <td>Houston</td>
          </tr>
          <tr>
            <td>Florida</td>
            <td>Miami</td>
          </tr>
          <tr>
            <td>Illinois</td>
            <td>Chicago</td>
          </tr>
        </tbody>
      </table>
    """),
    ui.br(),
    # Updated button using ui.input_action_button
    ui.input_action_button("complex_show", "Show Complex Result"),
    ui.div(ui.HTML("<div id='complex_result'>Result will appear here</div>"))
)

def server(input, output, session):
    # ----------------------------------------------------------------------
    # Create reactive calculations to hold the currently selected lookup values.
    # If no cell has yet been clicked, we default to "California".
    # ----------------------------------------------------------------------
    @reactive.Calc
    def basic_lookup():
        # For the basic example, we use the value from the clicked cell in the basic table.
        sel = input.basic_selected  # This will be a dict with keys row, col, value
        if sel is not None and "value" in sel:
            return sel["value"]
        return "California"
    
    @reactive.Calc
    def sheet1_lookup():
        sel = input.sheet1_selected
        if sel is not None and "value" in sel:
            return sel["value"]
        return "California"
    
    # ----------------------------------------------------------------------
    # Update the displayed formula (with colored parameters) for the basic example.
    # ----------------------------------------------------------------------
    @output
    @render.ui
    def basic_formula():
        lookup_val = basic_lookup()
        # Build the formula with inline styles for color
        # Here: B:B (blue), the lookup value (green), A:A (orange), and 0 (red)
        formula_html = (
            "INDEX(" +
            "<span style='color: blue;'>B:B</span>, " +
            "MATCH(" +
            "<span style='color: green;'>\"" + lookup_val + "\"</span>, " +
            "<span style='color: orange;'>A:A</span>, " +
            "<span style='color: red;'>0</span>" +
            "))"
        )
        return ui.HTML(formula_html)
    
    # ----------------------------------------------------------------------
    # Update the displayed formula for the complex example.
    # ----------------------------------------------------------------------
    @output
    @render.ui
    def complex_formula():
        lookup_val = sheet1_lookup()
        # In this formula, the lookup value comes from Sheet 1 (A2)
        # and the ranges for INDEX and MATCH are fixed for Sheet 2.
        formula_html = (
            "INDEX(" +
            "<span style='color: blue;'>[Sheet 2]!B:B</span>, " +
            "MATCH(" +
            "<span style='color: green;'>A2 (" + lookup_val + ")</span>, " +
            "<span style='color: orange;'>[Sheet 2]!C:C</span>, " +
            "<span style='color: red;'>0</span>" +
            "))"
        )
        return ui.HTML(formula_html)
    
    # ----------------------------------------------------------------------
    # Compute and display the result for the basic example when the button is clicked.
    # ----------------------------------------------------------------------
    @output
    @render.ui
    def basic_result():
        # The action button (basic_show) increments a counter each time it is clicked.
        # We use this as a trigger.
        if input.basic_show() > 0:
            lookup_val = basic_lookup()
            result = "Not found"
            for row in basic_data:
                if row["A"] == lookup_val:
                    result = row["B"]
                    break
            return ui.HTML(f"<p>Result: {result}</p>")
        else:
            return ui.HTML("<p>Result will appear here</p>")
    
    # ----------------------------------------------------------------------
    # Compute and display the result for the complex example.
    # ----------------------------------------------------------------------
    @output
    @render.ui
    def complex_result():
        if input.complex_show() > 0:
            lookup_val = sheet1_lookup()
            result = "Not found"
            for row in sheet2_data:
                if row["C"] == lookup_val:
                    result = row["B"]
                    break
            return ui.HTML(f"<p>Result: {result}</p>")
        else:
            return ui.HTML("<p>Result will appear here</p>")

app = App(app_ui, server)
