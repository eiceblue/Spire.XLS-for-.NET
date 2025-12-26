# spire.xls csharp excel creation
## create an excel workbook with five sheets and populate with data
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Create five empty sheets in the workbook
workbook.CreateEmptySheets(5);

// Iterate over each sheet in the workbook
for (int i = 0; i < 5; i++)
{
    // Get the current sheet
    Worksheet sheet = workbook.Worksheets[i];

    // Set the name of the sheet using the index
    sheet.Name = "Sheet" + i.ToString();

    // Populate the sheet with data in a grid-like pattern
    for (int row = 1; row <= 150; row++)
    {
        for (int col = 1; col <= 50; col++)
        {
            // Set the text in each cell of the sheet using the row and column numbers
            sheet.Range[row, col].Text = "row" + row.ToString() + " col" + col.ToString();
        }
    }
}
```

---

# spire.xls csharp create excel
## Create an Excel file with a single sheet and populate it with data
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Create an empty sheet in the workbook
workbook.CreateEmptySheets(1);

// Get the reference to the first sheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Populate the sheet with data
for (int row = 1; row <= 10000; row++)
{
    for (int col = 1; col <= 30; col++)
    {
        // Set the cell value to a combination of the current row and column numbers
        sheet.Range[row, col].Text = row.ToString() + "," + col.ToString();
    }
}

// Specify the filename for the resulting Excel file
String result = "CreateAnExcelWithOneSheet_result.xlsx";

// Save the workbook to the specified file in Excel 2010 format
workbook.SaveToFile(result, ExcelVersion.Version2010);

// Dispose of the workbook object
workbook.Dispose();
```

---

# spire.xls csharp create multiple excel files
## Create 50 Excel files with 5 sheets each and populate with data
```csharp
// Create 50 workbooks with 5 sheets each
for (int n = 0; n < 50; n++)
{
    // Create a new workbook object
    Workbook workbook = new Workbook();

    // Create 5 empty sheets in the workbook
    workbook.CreateEmptySheets(5);

    // Iterate through each sheet in the workbook
    for (int i = 0; i < 5; i++)
    {
        // Get the reference to the current sheet
        Worksheet sheet = workbook.Worksheets[i];

        // Set the name of the sheet based on the index
        sheet.Name = "Sheet" + i.ToString();

        // Populate the sheet with data
        for (int row = 1; row <= 150; row++)
        {
            for (int col = 1; col <= 50; col++)
            {
                // Set the cell value to a combination of the current row and column numbers
                sheet.Range[row, col].Text = "row" + row.ToString() + " col" + col.ToString();
            }
        }
    }

    // Specify the filename for the resulting Excel file, using the iteration number
    workbook.SaveToFile("Workbook" + n + ".xlsx", ExcelVersion.Version2010);

    // Dispose of the workbook object
    workbook.Dispose();
}
```

---

# spire.xls csharp hello world
## demonstrates basic excel file creation with text input and column auto-sizing
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the reference to the first sheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Set the cell value of cell A1 to "Hello World"
sheet.Range["A1"].Text = "Hello World";

// Auto-fit the columns to adjust their width based on the content
sheet.Range["A1"].AutoFitColumns();
```

---

# spire.xls csharp open existing file
## Load an existing Excel file and modify its content
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Load an existing Excel file from the specified path
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\templateAz2.xlsx");

// Add a new sheet with the name "MySheet"
Worksheet sheet = workbook.Worksheets.Add("MySheet");

// Set the value of cell A1 to "Hello World"
sheet.Range["A1"].Text = "Hello World";
```

---

# spire.xls csharp add label control
## Add a label control to Excel worksheet and set its text content
```csharp
// Add a label control to the worksheet
ILabelShape label = sheet.LabelShapes.AddLabel(10, 2, 30, 200);

// Set the text content of the label control
label.Text = "This is a Label Control";
```

---

# spire.xls csharp listbox
## add listbox control to excel worksheet
```csharp
// Add a listbox control to the worksheet
IListBox listBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80);
// Set the selection type to single (allows only one item to be selected)
listBox.SelectionType = SelectionType.Single;
// Set the initially selected index in the listbox
listBox.SelectedIndex = 2;
// Enable 3D shading for the listbox
listBox.Display3DShading = true;
// Specify the range to populate the listbox with data
listBox.ListFillRange = sheet.Range["A7:A12"];
```

---

# spire.xls csharp scrollbar control
## add scroll bar control to excel worksheet
```csharp
// Set a value and formatting for range B10
sheet.Range["B10"].Value2 = 1;
sheet.Range["B10"].Style.Font.IsBold = true;

// Add a scroll bar control to the worksheet
IScrollBarShape scrollBar = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20);
// Link the scroll bar control to cell B10
scrollBar.LinkedCell = sheet.Range["B10"];
// Set the minimum value of the scroll bar
scrollBar.Min = 1;
// Set the maximum value of the scroll bar
scrollBar.Max = 150;
// Set the incremental change when scrolling
scrollBar.IncrementalChange = 1;
// Enable 3D shading for the scroll bar
scrollBar.Display3DShading = true;
```

---

# spire.xls csharp add table with filter
## add a table with filter functionality to an Excel worksheet
```csharp
//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Create a List Object named in Table.
sheet.ListObjects.Create("Table", sheet.Range[1, 1, sheet.LastRow, sheet.LastColumn]);

//Set the BuiltInTableStyle for List object.
sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleLight9;
```

---

# Spire.XLS C# Table with Total Row
## Add total row to Excel table and configure sum calculations
```csharp
// Create a table with the data from the specified cell range
IListObject table = sheet.ListObjects.Create("Table", sheet.Range["A1:D4"]);

// Display the total row in the table
table.DisplayTotalRow = true;

// Add a total row to the table
table.Columns[0].TotalsRowLabel = "Total";
// Calculate the sum for column 1 in the total row
table.Columns[1].TotalsCalculation = ExcelTotalsCalculation.Sum;
// Calculate the sum for column 2 in the total row
table.Columns[2].TotalsCalculation = ExcelTotalsCalculation.Sum;
// Calculate the sum for column 3 in the total row
table.Columns[3].TotalsCalculation = ExcelTotalsCalculation.Sum;
```

---

# spire.xls csharp formatting
## apply subscript and superscript to Excel cells
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Set the text for cell B2
sheet.Range["B2"].Text = "This is an example of Subscript:";
// Set the text for cell D2
sheet.Range["D2"].Text = "This is an example of Superscript:"; 

// Set the RTF value of cell "B3" to "R100-0.06"
CellRange range = sheet.Range["B3"];
range.RichText.Text = "R100-0.06";

// Create a font and set the IsSubscript property to true
ExcelFont font = workbook.CreateFont();
font.IsSubscript = true;
font.Color = Color.Green;

// Set the font for the specified range of text in cell "B3"
range.RichText.SetFont(4, 8, font);

// Set the RichText value of cell "D3" to "a2 + b2 = c2"
range = sheet.Range["D3"];
range.RichText.Text = "a2 + b2 = c2";

// Create a font and set the IsSuperscript property to true
font = workbook.CreateFont();
font.IsSuperscript = true;

// Set the font for the specified range of text in cell "D3"
range.RichText.SetFont(1, 1, font);
range.RichText.SetFont(6, 6, font);
range.RichText.SetFont(11, 11, font);

// Auto-fit the columns to adjust their widths
sheet.AllocatedRange.AutoFitColumns();
```

---

# spire.xls csharp clone font style
## clone Excel font style and apply to cells
```csharp
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Add the text to the Excel sheet cell range A1.
sheet.Range["A1"].Text = "Text1";

//Set A1 cell range's CellStyle.
CellStyle style = workbook.Styles.Add("style");
style.Font.FontName = "Calibri";
style.Font.Color = Color.Red;
style.Font.Size = 12;
style.Font.IsBold = true;
style.Font.IsItalic = true;
sheet.Range["A1"].CellStyleName = style.Name;

//Clone the same style for B2 cell range.
CellStyle csOrieign = style.clone();
sheet.Range["B2"].Text = "Text2";
sheet.Range["B2"].CellStyleName = csOrieign.Name;

//Clone the same style for C3 cell range and then reset the font color for the text.
CellStyle csGreen = style.clone();
csGreen.Font.Color = Color.Green;
sheet.Range["C3"].Text = "Text3";
sheet.Range["C3"].CellStyleName = csGreen.Name;
```

---

# spire.xls csharp copy range
## copy cell range from source to destination in excel worksheet
```csharp
// Get the first worksheet
Worksheet sheet1 = workbook.Worksheets[0];

// Specify a destination range 
CellRange cells = sheet1.Range["G1:H19"];

// Copy the selected range to destination range 
sheet1.Range["B1:C19"].Copy(cells);
```

---

# spire.xls csharp copy data with style
## copy cell range with style from source to destination in Excel
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Get the default first worksheet
Worksheet worksheet = workbook.Worksheets[0];

//Set the values for some cells.
CellRange cells = worksheet.Range["A1:J50"];
for (int i = 1; i <= 10; i++)
{
    for (int j = 1; j <= 8; j++)
    {
        string text = string.Format((i - 1).ToString() + "," + (j - 1).ToString());
        cells[i, j].Text = text;
    }
}

//Get a source range (A1:D3).
CellRange srcRange = worksheet.Range["A1:D3"];

//Create a style object.
CellStyle style = workbook.Styles.Add("style");

//Specify the font attribute.
style.Font.FontName = "Calibri";

//Specify the shading color.
style.Font.Color = Color.Red;

//Specify the border attributes.
style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
style.Borders[BordersLineType.EdgeTop].Color = Color.Blue;
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
style.Borders[BordersLineType.EdgeBottom].Color = Color.Blue;
style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
style.Borders[BordersLineType.EdgeTop].Color = Color.Blue;
style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
style.Borders[BordersLineType.EdgeRight].Color = Color.Blue;
srcRange.CellStyleName = style.Name;

//Set the destination range
CellRange destRange = worksheet.Range["A12:D14"];

//Copy the range data with style
srcRange.Copy(destRange, true, true);
```

---

# spire.xls csharp copy formula values
## copy only formula values from one range to another
```csharp
// Set the copy option to only copy the formula values
CopyRangeOptions copyOptions = CopyRangeOptions.OnlyCopyFormulaValue;

// Copy a range of cells from A2:C2 to A5:C5 using the specified copy options
sheet.Copy(sheet.Range["A2:C2"], sheet.Range["A5:C5"], copyOptions);
```

---

# Spire.XLS C# Nested Grouping
## Create nested groups in Excel worksheet
```csharp
// Create a new workbook.
Workbook workbook = new Workbook();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Set the option to display summary rows above detail rows.
sheet.PageSetup.IsSummaryRowBelow = false;

// Insert sample data into cells.
sheet.Range["A1"].Value = "Project plan for project X";

// Summary row
sheet.Range["A3"].Value = "Set up"; 
sheet.Range["A4"].Value = "Task 1"; 
sheet.Range["A5"].Value = "Task 2"; 

sheet.Range["A7"].Value = "Launch";
sheet.Range["A8"].Value = "Task 1"; 
sheet.Range["A9"].Value = "Task 2"; 

// Group the rows that need to be grouped.
sheet.GroupByRows(2, 9, false); 
sheet.GroupByRows(4, 5, false); 
sheet.GroupByRows(8, 9, false); 
```

---

# spire.xls csharp table creation
## create Excel table with style
```csharp
// Get the first worksheet of the workbook
Worksheet sheet = workbook.Worksheets[0];  

// Add a new List Object to the worksheet with the name "table" and range [1, 1, 19, 5]
sheet.ListObjects.Create("table", sheet.Range[1, 1, 19, 5]);

// Apply a default style (TableStyleLight9) to the created table
sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleLight9;
```

---

# spire.xls csharp custom sort
## implement custom sorting in Excel spreadsheet
```csharp
// Set header to participate in sorting
workbook.DataSorter.IsIncludeTitle = false;
// Custom sort
workbook.DataSorter.SortColumns.Add(0, new String[]
    {"DD","CC", "BB", "AA", "HH","GG","FF","EE"});
workbook.DataSorter.Sort(workbook.Worksheets[0].Range["A1:A8"]);
```

---

# Spire.XLS C# Data Export
## Export Excel worksheet data to a data grid
```csharp
// Create a new workbook object to work with Excel files
Workbook workbook = new Workbook();

// Load file
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataExport.xlsx");

// Get first sheet
Worksheet sheet = workbook.Worksheets[0];

// Export data 
this.dataGrid1.DataSource = sheet.ExportDataTable();

// Dispose of the workbook object to free up resources
workbook.Dispose();
```

---

# spire.xls csharp data import
## import data from datatable to excel with styling
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Import data to data table
sheet.InsertDataTable(dataTable, true, 1, 1, -1, -1);

// Set body style
CellStyle oddStyle = workbook.Styles.Add("oddStyle");
oddStyle.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
oddStyle.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
oddStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
oddStyle.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
oddStyle.KnownColor = ExcelColors.LightGreen1;

CellStyle evenStyle = workbook.Styles.Add("evenStyle");
evenStyle.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
evenStyle.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
evenStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
evenStyle.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
evenStyle.KnownColor = ExcelColors.LightTurquoise;

foreach (CellRange range in sheet.AllocatedRange.Rows)
{
    if (range.Row % 2 == 0)
        range.CellStyleName = evenStyle.Name;
    else
        range.CellStyleName = oddStyle.Name;
}

// Set header style
CellStyle styleHeader = sheet.Rows[0].Style;
styleHeader.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
styleHeader.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
styleHeader.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
styleHeader.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
styleHeader.VerticalAlignment = VerticalAlignType.Center;
styleHeader.KnownColor = ExcelColors.Green;
styleHeader.Font.KnownColor = ExcelColors.White;
styleHeader.Font.IsBold = true;

// Auto-fit the columns to adjust their widths
sheet.AllocatedRange.AutoFitColumns();
// Auto-fit the rows to adjust their height
sheet.AllocatedRange.AutoFitRows();

// Set row height
sheet.Rows[0].RowHeight = 20;
```

---

# spire.xls csharp data sorting
## sort excel data by multiple columns
```csharp
// Get the first worksheet from the workbook
Worksheet worksheet = workbook.Worksheets[0];

// Add a sorting column: column 2 in ascending order
workbook.DataSorter.SortColumns.Add(2, OrderBy.Ascending);

// Add another sorting column: column 3  in ascending order
workbook.DataSorter.SortColumns.Add(3, OrderBy.Ascending);

// Perform the sorting operation on the specified range: A1 to E19
workbook.DataSorter.Sort(worksheet["A1:E19"]);
```

---

# Spire.XLS CSharp Group Management
## Expand and collapse grouped rows in Excel worksheet
```csharp
//Expand the grouped rows with ExpandCollapseFlags set to expand parent
sheet.Range["A16:G19"].ExpandGroup(GroupByType.ByRows, ExpandCollapseFlags.ExpandParent);

//Collapse the grouped rows
sheet.Range["A10:G12"].CollapseGroup(GroupByType.ByRows);
```

---

# Spire.XLS C# Export Data Format
## Export data from Excel worksheet with format options
```csharp
// Export DataTable without keeping data format
ExportTableOptions options = new ExportTableOptions();
options.KeepDataFormat = false;
options.RenameStrategy = RenameStrategy.Digit;

// Export data to data table
DataTable table = sheet.ExportDataTable(1, 1, sheet.LastDataRow, sheet.LastDataColumn, options);
```

---

# spire.xls csharp find and replace
## find and replace data in excel cells
```csharp
//Find the "Area" string
CellRange[] ranges = worksheet.FindAllString("Area", false, false);

//Traverse the found ranges
foreach (CellRange range in ranges)
{
    //Replace it with "Area Code"
    range.Text = "Area Code";
    //Highlight the color
    range.Style.Color = Color.Yellow;
}
```

---

# spire.xls csharp find data
## find text and number in specific range
```csharp
//Specify a range in the worksheet
CellRange range = sheet.Range[1, 1, 12, 8];

//Find all cells containing the specified text in the range
CellRange[] textRanges = range.FindAllString("E-iceblue", false, false);

//Iterate through the found text cells
if (textRanges.Length != 0)
{
    foreach (CellRange r in textRanges)
    {
        string address = r.RangeAddress;
        // Process the found cell address
    }
}

//Find all cells containing the specified number in the range
CellRange[] numberRanges = range.FindAllNumber(100, true);

//Iterate through the found number cells
if (numberRanges.Length != 0)
{
    foreach (CellRange r in numberRanges)
    {
        string address = r.RangeAddress;
        // Process the found cell address
    }
}
```

---

# spire.xls csharp find
## find string and number in excel worksheet
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.LoadFromFile("FindCellsSample.xlsx");

//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Find cells with the input string
CellRange[] textRanges = sheet.FindAllString("E-iceblue", false, false);

//Create a string builder
StringBuilder builder = new StringBuilder();

//Append the address of found cells in builder
if (textRanges.Length != 0)
{
    foreach (CellRange range in textRanges)
    {
        string address = range.RangeAddress;
        builder.AppendLine("The address of found text cell is: " + address);
    }
}
else
{
    builder.AppendLine("No cells that contain the text");
}

//Find cells with the input integer or double
CellRange[] numberRanges = sheet.FindAllNumber(100, true);

//Append the address of found cells in builder
if (numberRanges.Length != 0)
{
    foreach (CellRange range in numberRanges)
    {
        string address = range.RangeAddress;
        builder.AppendLine("The address of found number cell is: " + address);
    }
}
else
{
    builder.AppendLine("No cells that contain the number");
}
```

---

# spire.xls csharp find text
## find text in excel using regex
```csharp
// Get the first sheet
Worksheet worksheet = workbook.Worksheets[0];

// Find cell ranges by Regex
CellRange[] ranges = worksheet.FindAllString(".*North.", false, false, true);
string information = "";

// Get the information of every cell range
foreach (CellRange range in ranges)
{
    information += "RangeAddressLocal:" + range.RangeAddressLocal + "\r\n";
    information += "Text:" + range.Text + "\r\n";
}
```

---

# spire.xls csharp find text
## find text in cell range
```csharp
// Define the range to search for the text
CellRange range = sheet.Range["A16:B20"];

// Find all occurrences of the specified text in the range
CellRange[] resultRange = range.FindAll("e-iceblue1", FindType.Text, ExcelFindOptions.MatchEntireCellContent | ExcelFindOptions.MatchCase);

// Check if any occurrences were found
if (resultRange.Length != 0)
{
    // Iterate through the found ranges and append their addresses to the StringBuilder
    foreach (CellRange r in resultRange)
    {
        string address = r.RangeAddress;
        builder.AppendLine("In the range 'A16:B20', the address of the cell containing 'e-iceblue1' is: " + address);
    }
}
```

---

# spire.xls csharp table formatting
## format excel table with styles and total row
```csharp
// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Add a default table style to the table in the worksheet
sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9;

// Show total row for the table
sheet.ListObjects[0].DisplayTotalRow = true;

// Set calculation type
sheet.ListObjects[0].Columns[0].TotalsRowLabel = "Total";
sheet.ListObjects[0].Columns[1].TotalsCalculation = ExcelTotalsCalculation.None;
sheet.ListObjects[0].Columns[2].TotalsCalculation = ExcelTotalsCalculation.None;
sheet.ListObjects[0].Columns[3].TotalsCalculation = ExcelTotalsCalculation.Sum;
sheet.ListObjects[0].Columns[4].TotalsCalculation = ExcelTotalsCalculation.Sum;

// Show row stripes and column stripes using table style
sheet.ListObjects[0].ShowTableStyleRowStripes = true;
sheet.ListObjects[0].ShowTableStyleColumnStripes = true;
```

---

# spire.xls csharp goalseek
## implement goal seek functionality in excel
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Set value for cell "A1"
sheet.Range["A1"].Value = "100";

// Set formula for cell "A2"
CellRange targetCell = sheet.Range["A2"];
targetCell.Formula = "=SUM(A1+B1)";

// Variable cell
CellRange gussCell = sheet.Range["B1"];
Spire.Xls.GoalSeek goalSeek = new Spire.Xls.GoalSeek();

// Trial solution
GoalSeekResult result = goalSeek.TryCalculate(targetCell, 500, gussCell);

// Determine the solution
result.Determine();
```

---

# spire.xls csharp import arraylist
## Import data from ArrayList into Excel worksheet
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Create an empty worksheet
workbook.CreateEmptySheets(1);

//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Create an ArrayList object
ArrayList list = new ArrayList();

//Add strings in list
list.Add("Spire.Doc for .NET");
list.Add("Spire.XLS for .NET");
list.Add("Spire.PDF for .NET");
list.Add("Spire.Presentation for .NET");

//Insert array list in worksheet 
sheet.InsertArrayList(list, 1, 1, true);
```

---

# spire.xls csharp import data
## Import data from DataColumn to Excel worksheet
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Create an empty worksheet
workbook.CreateEmptySheets(1);

//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Create a DataTable object 
DataTable dataTable = new DataTable("Customer");
dataTable.Columns.Add("No", typeof(Int32));
dataTable.Columns.Add("Name", typeof(string));
dataTable.Columns.Add("City", typeof(string));

//Import the two columns of the data table to worksheet
DataColumn[] columns=new DataColumn[2]{dataTable.Columns[1],dataTable.Columns[2]};
sheet.InsertDataColumns(columns, true, 1, 1);
```

---

# Spire.XLS C# Import Data
## Import data from DataTable to Excel worksheet
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Create an empty worksheet
workbook.CreateEmptySheets(1);

//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Create a DataTable object 
DataTable dataTable = new DataTable("Customer");
dataTable.Columns.Add("No", typeof(Int32));
dataTable.Columns.Add("Name", typeof(string));
dataTable.Columns.Add("City", typeof(string));

//Import datatable in worksheet
sheet.InsertDataTable(dataTable, true, 1, 1);
```

---

# Spire.XLS C# Import Data from DataView
## Import data from a DataView into an Excel worksheet
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Create an empty worksheet
workbook.CreateEmptySheets(1);

//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Create a DataTable object 
DataTable dataTable = new DataTable("Customer");
dataTable.Columns.Add("No", typeof(Int32));
dataTable.Columns.Add("Name", typeof(string));
dataTable.Columns.Add("City", typeof(string));

//Create rows and add data
DataRow dr = dataTable.NewRow();
dr[0] = 1;
dr[1] = "Tom";
dr[2] = "New York";
dataTable.Rows.Add(dr);

//Import the data view of data table to worksheet
sheet.InsertDataView(dataTable.DefaultView, true, 1, 1);
```

---

# Spire.XLS C# Excel Controls
## Insert various controls (textbox, checkbox, radio button, combobox) into Excel worksheet
```csharp
// Add a textbox at position (9, 2) with width 25 and height 100
ITextBoxShape textbox = ws.TextBoxes.AddTextBox(9, 2, 25, 100);
// Set the text for the textbox
textbox.Text = "Hello World"; 

// Add a checkbox at position (11, 2) with width 15 and height 100
ICheckBox cb = ws.CheckBoxes.AddCheckBox(11, 2, 15, 100);
// Set the checkbox state to checked
cb.CheckState = Spire.Xls.CheckState.Checked;
// Set the text for the checkbox
cb.Text = "Check Box 1"; 

// Add a radio button at position (13, 2) with width 15 and height 100
IRadioButton rb = ws.RadioButtons.Add(13, 2, 15, 100);
// Set the text for the radio button
rb.Text = "Option 1"; 

// Add a combobox at position (15, 2) with width 15 and height 100
IComboBoxShape cbx = ws.ComboBoxes.AddComboBox(15, 2, 15, 100) as IComboBoxShape;
// Set the range of options for the combobox
cbx.ListFillRange = ws.Range["A41:A47"]; 
```

---

# spire.xls csharp html
## insert html string into excel cell
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Insert html code to cell "A1"
String htmlCode = "<div>first line<br>second line<br>third line</div>";
CellRange range = sheet["A1"];
range.HtmlString = htmlCode;
```

---

# spire.xls csharp replace and highlight
## Replace text and highlight cells in Excel
```csharp
// Find all occurrences of the string "Total" in the worksheet, including case-sensitive and whole word matches
CellRange[] ranges = worksheet.FindAllString("Total", true, true);

// Iterate through each found range
foreach (CellRange range in ranges)
{
    // Reset the text in the range by replacing it with "Sum"
    range.Text = "Sum";

    // Set the color of the range to yellow
    range.Style.Color = Color.Yellow;
}
```

---

# spire.xls csharp replace font
## replace font style in excel
```csharp
// Define the new style
CellStyle newStyle = workbook.Styles.Add("newStyle");
newStyle.Font.FontName = "Arial Black";
newStyle.Font.Size = 14;

// The old style which need to be replaced
CellStyle oldStyle = null;

for (int i = 0; i < workbook.Styles.Count; i++)
{
    if (workbook.Styles[i].Font.FontName == "Aleo")
    {
       oldStyle = sheet.Range["D9"].Style;            
    }
}

// Replace style
sheet.ReplaceAll("North America", oldStyle, "America", newStyle);
```

---

# Spire.XLS C# Text Replacement
## Replace partial text in Excel cells
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Set value for cell "A1"
sheet.Range["A1"].Text = "Hello World";

// Automatically adjusting the column width to fit the content.
sheet.Range["A1"].AutoFitColumns();

// Replace Partial Text
sheet.CellList[0].TextPartReplace("World", "Spire");
```

---

# Spire.XLS C# Data Retrieval
## Extract specific rows from Excel file based on cell value
```csharp
// Create a new workbook instance
Workbook newBook = new Workbook();

// Get the first worksheet.
Worksheet newSheet = newBook.Worksheets[0];

// Create a new workbook instance and load the sample Excel file.
Workbook workbook = new Workbook();
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_3.xlsx");

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Retrieve data and extract it to the first worksheet of the new excel workbook.
int i = 1;
int columnCount = sheet.Columns.Length;
foreach (CellRange range in sheet.Columns[0])
{
    if (range.Text == "teacher")
    {
        CellRange sourceRange = sheet.Range[range.Row, 1, range.Row, columnCount];
        CellRange destRange = newSheet.Range[i, 1, i, columnCount];
        sheet.Copy(sourceRange, destRange, true);
        i++;
    }
}
```

---

# spire.xls csharp set array values
## insert array data into excel range
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Create an empty worksheet
workbook.CreateEmptySheets(1);

// Get the worksheet
Worksheet sheet = workbook.Worksheets[0];

// Set the value of max row and column
int maxRow = 10000;
int maxCol = 200;

// Output an array of data to a range of worksheet
object[,] myarray = new object[maxRow + 1, maxCol + 1];
for (int i = 0; i <= maxRow; i++)
    for (int j = 0; j <= maxCol; j++)
    {
        myarray[i, j] = i + j;
    }

// Insert the array of data into the worksheet starting from cell (1, 1)
sheet.InsertArray(myarray, 1, 1);
```

---

# spire.xls csharp show formula and result
## Export Excel data as DataTable with option to show formulas or results
```csharp
//Show formula
DataTable dt = sheet.ExportDataTable(sheet.AllocatedRange, false, false);
//Show in DataGridView
this.dataGridView1.DataSource = dt;
```

```csharp
//Show result
DataTable dt = sheet.ExportDataTable(sheet.AllocatedRange, false, true);
//Show in DataGridView
this.dataGridView1.DataSource = dt;
```

---

# Spire.XLS C# Data Processing
## Split Excel data into multiple columns based on space delimiter
```csharp
//Split data into separate columns by the delimited characters – space.
string[] splitText = null;
string text = null;
for (int i = 1; i < sheet.LastRow; i++)
{
    text = sheet.Range[i + 1, 1].Text;
    splitText = text.Split(' ');
    for (int j = 0; j < splitText.Length; j++)
    {
        sheet.Range[i + 1, 1 + j + 1].Text = splitText[j];
    }
}
```

---

# spire.xls csharp subtotal
## create subtotals in excel worksheet
```csharp
// Select the range of data to be used for subtotals (in this case, columns A and B, rows 1 to 18)
CellRange range = sheet.Range["A1:B18"];
// Apply subtotals to the selected data using the "Sum" function
sheet.Subtotal(range, 0, new int[] {1}, SubtotalTypes.Sum, true, false, true);
```

---

# spire.xls csharp richtext
## write rich text to excel cells with different font styles
```csharp
// Create font styles for different formatting options
ExcelFont fontBold = workbook.CreateFont();
fontBold.IsBold = true;

ExcelFont fontUnderline = workbook.CreateFont();
fontUnderline.Underline = FontUnderlineType.Single;

ExcelFont fontItalic = workbook.CreateFont();
fontItalic.IsItalic = true;

ExcelFont fontColor = workbook.CreateFont();
fontColor.KnownColor = ExcelColors.Green;

// Get the rich text object for cell B11 in the worksheet
RichText richText = sheet.Range["B11"].RichText;

// Set the text content for the rich text
richText.Text = "Bold and underlined and italic and colored text.";

// Apply different font styles to specific parts of the rich text
richText.SetFont(0, 3, fontBold); 
richText.SetFont(9, 18, fontUnderline);
richText.SetFont(24, 29, fontItalic); 
richText.SetFont(35, 41, fontColor);
```

---

# Spire.XLS C# Cell Access
## Access Excel cells using different methods
```csharp
//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Access cell by its name
CellRange range1 = sheet.Range["A1"];
builder.AppendLine("Value of range1: " + range1.Text);

//Access cell by index of row and column
CellRange range2 = sheet.Range[2,1];
builder.AppendLine("Value of range2: " + range2.Text);

//Access cell in cell collection
CellRange range3 = sheet.Cells[2];
builder.AppendLine("Value of range3: " + range3.Text);
```

---

# Spire.XLS C# Multiple Fonts in Cell
## Apply multiple fonts within a single Excel cell
```csharp
//Create a font object in workbook, setting the font color, size and type.
ExcelFont font1 = workbook.CreateFont();
font1.KnownColor = ExcelColors.LightBlue;
font1.IsBold = true;
font1.Size = 10;

//Create another font object specifying its properties.
ExcelFont font2 = workbook.CreateFont();
font2.KnownColor = ExcelColors.Red;
font2.IsBold = true;
font2.IsItalic = true;
font2.FontName = "Times New Roman";
font2.Size = 11;

//Write a RichText string to the cell 'A1', and set the font for it.
RichText richText = sheet.Range["H5"].RichText;
richText.Text = "This document was created with Spire.XLS for .NET.";
richText.SetFont(0, 29, font1);
richText.SetFont(31, 48, font2);
```

---

# spire.xls csharp style application
## apply custom style to used cells in excel worksheets
```csharp
// Create a new CellStyle object and name it "Mystyle".
CellStyle cellStyle = workbook.Styles.Add("Mystyle");

// Set the background color of the cell style to transparent.
cellStyle.Color = System.Drawing.Color.Transparent;

// Set the border color of the cell style to black.
cellStyle.Borders.KnownColor = ExcelColors.Black;

// Set the line style of the borders in the cell style to thin.
cellStyle.Borders.LineStyle = LineStyleType.Thin;

// Set the line style of the diagonal-down border to none.
cellStyle.Borders[BordersLineType.DiagonalDown].LineStyle = LineStyleType.None;

// Set the line style of the diagonal-up border to none.
cellStyle.Borders[BordersLineType.DiagonalUp].LineStyle = LineStyleType.None;

// Iterate through each worksheet in the workbook.
foreach (Worksheet worksheet in workbook.Worksheets)
{
    // Apply only the style to the used cells 
    worksheet.ApplyStyle(cellStyle, false, false);
}
```

---

# spire.xls csharp autofit
## AutoFit columns and rows based on cell value in Excel
```csharp
//Set value for B8
CellRange cell = worksheet.Range["B8"];
cell.Text = "Welcome to Spire.XLS!";

//Set the cell style
CellStyle style = cell.Style;
style.Font.Size = 16;
style.Font.IsBold = true;

//Autofit column width and row height based on cell value
cell.AutoFitColumns();
cell.AutoFitRows();
```

---

# spire.xls csharp text to number conversion
## convert text format to number format in excel cells
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the Excel document from disk
workbook.LoadFromFile("Sample.xlsx");

//Get the first worksheet
Worksheet worksheet = workbook.Worksheets[0];

//Convert text string format to number format
worksheet.Range["D2:D8"].ConvertToNumber();

// Save the workbook
workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2013);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp cell format
## copy cell format from one column to another
```csharp
//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Copy the cell format from column 2 and apply to cells of column 5.
int count = sheet.Rows.Length;
for (int i = 1; i < count + 1; i++)
{
    sheet.Range[string.Format("E{0}", i)].Style = sheet.Range[string.Format("B{0}", i)].Style;
}
```

---

# Spire.XLS C# Cell Counting
## Count the number of cells in an Excel worksheet
```csharp
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Get the number of cells.
int cellCount = sheet.Cells.Length;
```

---

# Spire.XLS C# Cut Cells
## Cut cells from one position to another in Excel worksheet
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Define the original source range to be copied (cells A1 to C5)
CellRange Ori = sheet.Range["A1:C5"];

// Define the destination range where the source range will be copied to (cells A26 to C30)
CellRange Dest = sheet.Range["A26:C30"];

// Copy the range to other position
sheet.Copy(Ori, Dest, true, true, true);

// Remove all content in original cells
foreach (CellRange cr in Ori)
{
    cr.ClearAll();
}
```

---

# spire.xls csharp merged cells
## detect and unmerge merged cells in excel worksheet
```csharp
//Get the merged cell ranges in the first worksheet and put them into a CellRange array.
CellRange[] range = sheet.MergedCells;

//Traverse through the array and unmerge the merged cells.
foreach (CellRange cell in range)
{
    cell.UnMerge();
}
```

---

# Spire.XLS C# Duplicate Cell Range
## Copy data from source range to destination range while maintaining formatting
```csharp
//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Copy data from source range to destination range and maintain the format.
sheet.Copy(sheet.Range["A6:F6"], sheet.Range["A16:F16"], true);
```

---

# spire.xls csharp empty cell
## remove content from excel cells using different methods
```csharp
//Set the value as null to remove the original content from the Excel Cell.
sheet.Range["C6"].Value = "";

//Clear the contents to remove the original content from the Excel Cell.
sheet.Range["B6"].ClearContents();

//Remove the contents with format from the Excel cell.
sheet.Range["D6"].ClearAll();
```

---

# spire.xls csharp filter
## filter cells by cell color
```csharp
//Create an auto filter in the sheet and specify the range to be filtered
sheet.AutoFilters.Range = sheet.Range["G1:G19"];

//Get the column to be filtered
FilterColumn filtercolumn = (FilterColumn)sheet.AutoFilters[0];

//Add a color filter to filter the column based on cell color
sheet.AutoFilters.AddFillColorFilter(filtercolumn, Color.Red);

//Filter the data.
sheet.AutoFilters.Filter();
```

---

# spire.xls csharp find cells by style
## Find cells with the same style name as a reference cell
```csharp
//Get the cell style name
string styleName = sheet.Range["A1"].CellStyleName;

CellRange ranges = sheet.AllocatedRange;
foreach (CellRange cc in ranges)
{
    //Find the cells which have the same style name
    if (cc.CellStyleName == styleName)
    {
        //Set value
        cc.Value = "Same style";
    }
}
```

---

# spire.xls csharp find formula cells
## find cells containing specific formula in excel worksheet
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Find the cells that contain formula "=SUM(A11,A12)"
CellRange[] ranges = sheet.FindAll("=SUM(A11,A12)", FindType.Formula, ExcelFindOptions.None);

// Create a string builder
StringBuilder builder = new StringBuilder();

// Append the address of found cells to builder
if (ranges.Length != 0)
{
    foreach (CellRange range in ranges)
    {
        string address = range.RangeAddress;
        builder.AppendLine("The address of found cell is: " + address);
    }
}
else
{
    builder.AppendLine("No cell contain the formula");
}
```

---

# spire.xls csharp cell region
## get and clear cell current region
```csharp
// Get the current region of the cell starting from cell A1 and clear its contents.
IXLSRange xlRange = sheet.Range["A1"].CurrentRegion;
foreach (CellRange range in xlRange)
{
    range.ClearAll();
}
```

---

# spire.xls csharp get cell address
## retrieve cell range addresses and properties from excel worksheet
```csharp
//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Get a cell range
CellRange range = sheet.Range["A1:B5"];

//Get address of range
string address = range.RangeAddressLocal;

//Get the cell count of range
int count = range.CellsCount;

//Get the address of the entire column of range
string entireColAddress = range.EntireColumn.RangeAddressLocal;

//Get the address of the entire row of range
string entireRowAddress = range.EntireRow.RangeAddressLocal;
```

---

# spire.xls csharp cell data type
## get cell data type and display in adjacent cells
```csharp
//Get the cell types of the cells in range "H2:H7"
foreach (CellRange range in sheet.Range["H2:H7"])
{
    XlsWorksheet.TRangeValueType cellType = sheet.GetCellType(range.Row, range.Column, false);
    sheet[range.Row, range.Column + 1].Text = cellType.ToString();
    sheet[range.Row, range.Column + 1].Style.Font.Color = Color.Red;
    sheet[range.Row, range.Column + 1].Style.Font.IsBold = true;
}
```

---

# spire.xls get cell displayed text
## get the actual displayed text of a cell in excel
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Get first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];

//Set value for B8
CellRange cell = worksheet.Range["B8"];
cell.NumberValue = 0.012345;

//Set the cell style
CellStyle style = cell.Style;
style.NumberFormat = "0.00";

//Get the cell value
string cellValue = cell.Value;

//Get the displayed text of the cell
string displayedText = cell.DisplayedText;
```

---

# Spire.XLS C# Get Cell Value by Name
## Retrieve cell value using cell name reference
```csharp
//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Specify a cell by its name.
CellRange cell = sheet.Range["A2"];

//Get value of cell "A2".
string cellValue = cell.Value;
```

---

# spire.xls csharp get intersection
## get intersection of two cell ranges
```csharp
// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Get the two ranges.
CellRange range = sheet.Range["A2:D7"].Intersect(sheet.Range["B2:E8"]);

StringBuilder content = new StringBuilder();
content.AppendLine("The intersection of the two ranges \"A2:D7\" and \"B2:E8\" is:");

// Get the intersection of the two ranges.
foreach (CellRange r in range)
{
    content.AppendLine(r.Value.ToString());
}
```

---

# spire.xls csharp hide cell content
## Hide cell content by setting number format
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Hide the area by setting the number format as ";;;"
sheet.Range["C5:D6"].NumberFormat = ";;;";
```

---

# Spire.XLS C# Merge Cells
## Merge columns and ranges in Excel worksheets
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Merge the seventh column in Excel file
workbook.Worksheets[0].Columns[6].Merge();

// Merge the particular range in Excel file
workbook.Worksheets[0].Range["A14:D14"].Merge();
```

---

# spire.xls csharp active selection range
## obtain information about the active selection range in excel worksheet
```csharp
// Get the first sheet
Worksheet worksheet = workbook.Worksheets[0];

string information = null;

// Get the information of the active selection range
foreach (CellRange range in worksheet.ActiveSelectionRange)
{
    information += "RangeAddressLocal:" + range.RangeAddressLocal + "\r\n";
    information += "ColumnCount:" + range.ColumnCount + "\r\n";
    information += "ColumnWidth:" + range.ColumnWidth + "\r\n";
    information += "Column:" + range.Column + "\r\n";
    information += "RowCount:" + range.RowCount + "\r\n";
    information += "RowHeight:" + range.RowHeight + "\r\n";
    information += "Row:" + range.Row + "\r\n";
}
```

---

# Spire.XLS C# Copy Formula Values
## Copy only formula values from one range to another in Excel
```csharp
// Set the copy option to only copy formula values
CopyRangeOptions copyOptions = CopyRangeOptions.OnlyCopyFormulaValue;

// Define the source range to be copied
CellRange sourceRange = sheet.Range["A6:E6"];

// Copy the source range to a destination range using the specified copy options
sheet.Copy(sourceRange, sheet.Range["A8:E8"], copyOptions);

// Copy the source range to another destination range using the same copy options
sourceRange.Copy(sheet.Range["A10:E10"], copyOptions);
```

---

# spire.xls csharp cell formatting
## set cell fill pattern and color
```csharp
// Set the cell color for range B7:F7 to yellow
worksheet.Range["B7:F7"].Style.Color = Color.Yellow;

// Set the cell fill pattern for range B8:F8 to 125% gray
worksheet.Range["B8:F8"].Style.FillPattern = ExcelPatternType.Percent125Gray;
```

---

# spire.xls csharp DBNum formatting
## Set DB number format for Excel cells
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Create an empty worksheet
workbook.CreateEmptySheets(1);

// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Set values for cells A1, A2, and A3
sheet.Range["A1"].Value2 = 123;
sheet.Range["A2"].Value2 = 456;
sheet.Range["A3"].Value2 = 789;

// Get the cell range A1:A3
CellRange range = sheet.Range["A1:A3"];

// Set the DB num format for the range
range.NumberFormat = "[DBNum2][$-804]General";

// Auto fit columns for the range
range.AutoFitColumns();
```

---

# spire.xls csharp shrink text
## shrink text to fit in a cell
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Define the cell range to shrink text
CellRange cell = sheet.Range["B13:C13"];

// Enable ShrinkToFit for the cell range
CellStyle style = cell.Style;
style.ShrinkToFit = true;
```

---

# Spire.XLS C# Traverse Cell Values
## Traverse through cells in a worksheet and retrieve their values
```csharp
// Get the first worksheet in the workbook
Worksheet worksheet = workbook.Worksheets[0];

// Get the collection of cell ranges in the worksheet
CellRange[] cellRangeCollection = worksheet.Cells;

// Traverse through the cells and retrieve their values
foreach (CellRange cellRange in cellRangeCollection)
{
    // Set the string format for displaying the cell address and value
    string result = string.Format("Cell: " + cellRange.RangeAddress + "   Value: " + cellRange.Value);
}
```

---

# spire.xls csharp ungroup cells
## ungroup Excel rows by specified range
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Ungroup rows 10 to 12
sheet.UngroupByRows(10, 12);

// Ungroup rows 16 to 19
sheet.UngroupByRows(16, 19);
```

---

# spire.xls csharp unmerge cells
## unmerge specific cells in an excel worksheet
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Unmerge the cells in range F2
sheet.Range["F2"].UnMerge();

// Unmerge the cells in range F7
sheet.Range["F7"].UnMerge();
```

---

# spire.xls csharp explicit line breaks
## demonstrates how to use explicit line breaks in Excel cells
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Get the first default worksheet in the workbook
Worksheet sheet1 = workbook.Worksheets[0];

// Specify a cell range
CellRange c5 = sheet1.Range["C5"];

// Set the cell width for the specified range
sheet1.SetColumnWidth(c5.Column, 70);

// Put the string value with explicit line breaks into the cell
c5.Value = "Spire.XLS for .NET is a professional Excel .NET API\n that can be used to create, read, \nwrite, convert and print Excel files in any type \nof .NET(C#, VB.NET, ASP.NET, .NET Core) application. \nSpire.XLS for .NET offers object model\n Excel API for speeding up Excel programming in .NET platform -\n create new Excel documents from template, edit existing \nExcel documents and \nconvert Excel files.";

// Enable text wrap for the cell
c5.IsWrapText = true;
```

---

# spire.xls csharp text formatting
## wrap or unwrap text in excel cells
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Wrap the text in cell C1
sheet.Range["C1"].Text = "e-iceblue is in facebook and welcome to like us";
sheet.Range["C1"].Style.WrapText = true;

// Wrap the text in cell D1
sheet.Range["D1"].Text = "e-iceblue is in twitter and welcome to follow us";
sheet.Range["D1"].Style.WrapText = true;

// Unwrap the text in cell C2
sheet.Range["C2"].Text = "http://www.facebook.com/pages/e-iceblue/139657096082266";
sheet.Range["C2"].Style.WrapText = false;

// Unwrap the text in cell D2
sheet.Range["D2"].Text = "https://twitter.com/eiceblue";
sheet.Range["D2"].Style.WrapText = false;

// Set the text color and size of Range["C1:D1"]
sheet.Range["C1:D1"].Style.Font.Size = 15;
sheet.Range["C1:D1"].Style.Font.Color = Color.Blue;

// Set the text color and size of Range["C2:D2"]
sheet.Range["C2:D2"].Style.Font.Size = 15;
sheet.Range["C2:D2"].Style.Font.Color = Color.DeepSkyBlue;
```

---

# spire.xls csharp column auto-fit
## auto-fit columns in specified range
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Autofit the column width for columns 2 to 5 in the worksheet
sheet.AutoFitColumn(2, 2, 5);
```

---

# spire.xls csharp autofit row
## auto fit row in specified range
```csharp
// Autofit the second row of the worksheet, excluding merged cells
sheet.AutoFitRow(2, 1, 2, false);
```

---

# Spire.XLS C# AutoFit Check
## Check if rows or columns have auto-fit settings in Excel
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Check if the second row has auto-fit row height set
bool isRowAutofit = workbook.Worksheets[0].GetRowIsAutoFit(2);
if (isRowAutofit)
{
    result.AppendLine("The second row is auto fit row height.");
}
else 
{
    result.AppendLine("The second row is not auto fit row height.");
}

// Check if the second column has auto-fit column width set
bool isColAutofit = workbook.Worksheets[0].GetColumnIsAutoFit(2);
if (isColAutofit)
{
    result.AppendLine("The second column is auto fit column width.");
}
else
{
    result.AppendLine("The second column is not auto fit column width.");
}
```

---

# spire.xls csharp check hidden row column
## Check if a row or column is hidden in an Excel worksheet
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Specify the row and column index to check
int rowIndex = 2;
int columnIndex = 2;

// Check if the second row is hidden
bool rowIsHide = sheet.GetRowIsHide(rowIndex);
if (rowIsHide)
{
    result.AppendLine("The second row is hidden.");
}
else
{
    result.AppendLine("The second row is not hidden.");
}

// Check if the second column is hidden
bool columnIsHide = sheet.GetColumnIsHide(columnIndex);
if (columnIsHide)
{
    result.AppendLine("The second column is hidden.");
}
else
{
    result.AppendLine("The second column is not hidden.");
}
```

---

# spire.xls csharp column operations
## copy columns within and between worksheets
```csharp
// Get the first worksheet in the workbook
Worksheet sheet1 = workbook.Worksheets[0];

// Get the second worksheet in the workbook
Worksheet sheet2 = workbook.Worksheets[1];

// Copy the first column (column index 0) to the third column (column index 2) in the same sheet
sheet1.Copy(sheet1.Columns[0], sheet1.Columns[2], true, true, true);

// Copy the first column (column index 0) to the second column (column index 1) in a different sheet
sheet1.Copy(sheet1.Columns[0], sheet2.Columns[1], true, true, true);
```

---

# spire.xls csharp copy rows
## copy rows within and between worksheets
```csharp
// Get the second worksheet in the workbook
Worksheet sheet1 = workbook.Worksheets[1];

// Get the first worksheet in the workbook
Worksheet sheet2 = workbook.Worksheets[0];

// Copy the first row (row index 0) to the third row (row index 2) in the same sheet
sheet1.Copy(sheet1.Rows[0], sheet1.Rows[2], true, true, true);

// Copy the first row (row index 0) to the second row (row index 1) in a different sheet
sheet1.Copy(sheet1.Rows[0], sheet2.Rows[1], true, true, true);
```

---

# Spire.XLS C# Copy Column and Row
## Demonstrates how to copy a single column and row to different locations in an Excel worksheet
```csharp
// Get the first worksheet in the workbook
Worksheet sheet1 = workbook.Worksheets[0];

// Specify the destination range to copy one column (column G)
CellRange columnCells = sheet1.Range["G1:G19"];

// Copy the second column (column index 1) to the destination range
sheet1.Columns[1].Copy(columnCells);

// Specify the destination range to copy one row (row 21, columns A to E)
CellRange rowCells = sheet1.Range["A21:E21"];

// Copy the first row (row index 0) to the destination range
sheet1.Rows[0].Copy(rowCells);
```

---

# Spire.XLS C# Copy with Options
## Copy a range from one worksheet to another with specific options
```csharp
// Get the first worksheet in the workbook
Worksheet sheet1 = workbook.Worksheets[0];

// Add a new worksheet as the destination sheet
Worksheet destinationSheet = workbook.Worksheets.Add("DestSheet");

// Specify the range to be copied from the original sheet (B2:D4)
CellRange cellRange = sheet1.Range["B2:D4"];

// Copy the specified range to the added worksheet, keeping the original styles and updating references
workbook.Worksheets[0].Copy(cellRange, workbook.Worksheets[1], 2, 1, true, true);
```

---

# spire.xls csharp delete blank rows and columns
## delete empty rows and columns from Excel worksheet
```csharp
// Delete blank rows from the worksheet
for (int i = sheet.Rows.Length - 1; i >= 0; i--)
{
    if (sheet.Rows[i].IsBlank)
    {
        sheet.DeleteRow(i + 1);
    }
}

// Delete blank columns from the worksheet
for (int j = sheet.Columns.Length - 1; j >= 0; j--)
{
    if (sheet.Columns[j].IsBlank)
    {
        sheet.DeleteColumn(j + 1);
    }
}
```

---

# spire.xls csharp delete rows and columns
## delete multiple rows and columns from an Excel worksheet
```csharp
// Delete 4 rows starting from the fifth row (rows 5, 6, 7, and 8)
sheet.DeleteRow(5, 4);

// Delete 2 columns starting from the second column (columns B and C)
sheet.DeleteColumn(2, 2);
```

---

# spire.xls csharp worksheet
## get default row and column count
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Clear all existing worksheets in the workbook
workbook.Worksheets.Clear();

// Create a new empty worksheet
Worksheet sheet = workbook.CreateEmptySheet();

// Get the default row count and column count of the worksheet
int rowCount = sheet.Rows.Length;
int columnCount = sheet.Columns.Length;
```

---

# spire.xls csharp row column grouping
## group rows and columns in excel worksheet
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Group rows 1 to 5 (excluding child groups)
sheet.GroupByRows(1, 5, false);

// Group columns 1 to 3 (excluding child groups)
sheet.GroupByColumns(1, 3, false);
```

---

# spire.xls csharp row column headers
## hide or show row and column headers in Excel worksheet
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Hide the headers of rows and columns
sheet.RowColumnHeadersVisible = false;

// Show the headers of rows and columns
// sheet.RowColumnHeadersVisible = true;
```

---

# Spire.XLS C# Hide Rows and Columns
## Demonstrate how to hide specific rows and columns in an Excel worksheet
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet worksheet = workbook.Worksheets[0];

// Hide the second column of the worksheet
worksheet.HideColumn(2);

// Hide the fourth row of the worksheet
worksheet.HideRow(4);
```

---

# spire.xls csharp rows columns
## insert rows and columns in excel worksheet
```csharp
// Insert a row into the worksheet at index 2
worksheet.InsertRow(2);

// Insert a column into the worksheet at index 2
worksheet.InsertColumn(2);

// Insert multiple rows into the worksheet starting at index 5, with a count of 2
worksheet.InsertRow(5, 2);

// Insert multiple columns into the worksheet starting at index 5, with a count of 2
worksheet.InsertColumn(5, 2);
```

---

# Spire.XLS C# Remove Row Based on Keyword
## Remove Excel row containing specific keyword
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Find the string "Address" in the worksheet
CellRange cr = sheet.FindString("Address", false, false);

// Delete the row that includes the found string
sheet.DeleteRow(cr.Row);
```

---

# Spire.XLS C# Column Width
## Set column width in pixels
```csharp
// Set the width of the third column to 400 pixels
sheet.SetColumnWidthInPixels(3, 400);
```

---

# spire.xls csharp column width
## set default column width for worksheet
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Set the default column width to 25 units
sheet.DefaultColumnWidth = 25;
```

---

# Spire.XLS C# Default Style
## Set default row and column styles in Excel
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Create a cell style and set the color to yellow
CellStyle style = workbook.Styles.Add("Mystyle");
style.Color = Color.Yellow;

// Set the default style for the first row using the created style
sheet.SetDefaultRowStyle(1, style);

// Set the default style for the first column using the created style
sheet.SetDefaultColumnStyle(1, style);
```

---

# spire.xls csharp set row height
## set default row height for worksheet
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Set the default row height to 30 units
sheet.DefaultRowHeight = 30;
```

---

# spire.xls csharp row column dimensions
## set row height and column width in excel worksheet
```csharp
// Get the first worksheet in the workbook
Worksheet worksheet = workbook.Worksheets[0];

// Set the width of column 4 to 30 units
worksheet.SetColumnWidth(4, 30);

// Set the height of row 4 to 30 units
worksheet.SetRowHeight(4, 30);
```

---

# spire.xls csharp summary column direction
## set summary column to the right of grouped columns
```csharp
// Group columns 1 to 4
sheet.GroupByColumns(1, 4, true);

// Set the summary columns to the right of the details
sheet.PageSetup.IsSummaryColumnRight = true;
```

---

# spire.xls csharp summary row direction
## set summary row position in grouped Excel data
```csharp
// Group rows 1 to 4
sheet.GroupByRows(1, 4, true);

// Set the summary rows above the details
sheet.PageSetup.IsSummaryRowBelow = false;
```

---

# spire.xls csharp unhide rows and columns
## Unhide specific rows and columns in Excel worksheet
```csharp
// Unhide row 15
sheet.ShowRow(15);

// Unhide column 4
sheet.ShowColumn(4);
```

---

# Spire.XLS C# Picture Alignment
## Align picture within a cell in Excel
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Insert an image at the specific cell (1, 1)
ExcelPicture picture = sheet.Pictures.Add(1, 1, imagePath);

// Adjust the column width and row height so that the cell can contain the picture
sheet.Columns[0].ColumnWidth = 40;
sheet.Rows[0].RowHeight = 200;

// Set the horizontal offset of the image within the cell to 100
picture.LeftColumnOffset = 100;

// Set the vertical offset of the image within the cell to 25
picture.TopRowOffset = 25;
```

---

# spire.xls csharp image compression
## compress pictures in excel worksheets
```csharp
// Compress the picture quality for all pictures in all worksheets
foreach (Worksheet sheet in workbook.Worksheets)
{
    foreach (ExcelPicture picture in sheet.Pictures)
    {
        // Set the compression level to 50 (50% of original quality)
        picture.Compress(50);
    }
}
```

---

# spire.xls csharp image copy
## copy picture from one worksheet to another
```csharp
// Get the first worksheet in the workbook
Worksheet sheet1 = workbook.Worksheets[0];

// Add a new worksheet as the destination sheet
Worksheet destinationSheet = workbook.Worksheets.Add("DestSheet");

// Get the first picture from the first worksheet
ExcelPicture sourcePicture = sheet1.Pictures[0];

// Get the image from the picture
Image image = sourcePicture.Picture;

// Add the image into the added worksheet at cell (2, 2)
destinationSheet.Pictures.Add(2, 2, image);
```

---

# spire.xls csharp worksheet images
## delete all images from worksheet
```csharp
// Delete all images from the worksheet.
for (int i = sheet.Pictures.Count - 1; i >= 0; i--)
{
    sheet.Pictures[i].Remove();
}
```

---

# spire.xls csharp get image crop position
## extract position and dimensions of a cropped image in excel worksheet
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the Excel document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadImages.xlsx");

//Get the first worksheet
Worksheet sheet1 = workbook.Worksheets[0];

//Get the image from the first sheet
ExcelPicture picture = sheet1.Pictures[0];

//Get the cropped position
int left = picture.Left;
int top = picture.Top;
int width = picture.Width;
int height = picture.Height;
```

---

# spire.xls csharp get embedded images
## retrieve embedded images from excel worksheet
```csharp
// Access the first worksheet in the workbook
Worksheet sheet = wb.Worksheets[0];

// Retrieve an array of Excel pictures from the worksheet
ExcelPicture[] pc = sheet.CellImages;

// Iterate through each Excel picture in the array
for (int i = 0; i < pc.Length; i++)
{
    ExcelPicture ep = pc[i];
    Image image = ep.Picture;
}
```

---

# spire.xls csharp background image
## insert background image into excel worksheet
```csharp
//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Open an image. 
Bitmap bm = new Bitmap(Image.FromFile(imagePath));

//Set the image to be background image of the worksheet.
sheet.PageSetup.BackgoundImage = bm;
```

---

# spire.xls insert image in cell
## insert image into worksheet cell using Spire.XLS
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet worksheet = workbook.Worksheets[0];

// Embed an image into the first cell
worksheet.Cells[0].InsertOrUpdateCellImage("image_path.png", true);
```

---

# spire.xls csharp web image
## insert image from web URL into Excel worksheet
```csharp
// Create a new workbook.
Workbook workbook = new Workbook();

// Get the first sheet from the workbook.
Worksheet sheet = workbook.Worksheets[0];

// Specify the URL of the image to be downloaded.
string URL = "http://www.e-iceblue.com/downloads/demo/Logo.png";

// Instantiate a web client object.
WebClient webClient = new WebClient();

// Extract the image data into a memory stream.
MemoryStream objImage = new System.IO.MemoryStream(webClient.DownloadData(URL));

// Create an Image object from the memory stream.
Image image = Image.FromStream(objImage);

// Add the image to the worksheet at a specific location (row 3, column 2).
sheet.Pictures.Add(3, 2, image);
```

---

# spire.xls csharp image positioning
## locate and position images in excel worksheet
```csharp
//Create a Workbook
Workbook workbook = new Workbook();

//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Get the first picture from the sheet.
ExcelPicture pic = sheet.Pictures[0];

// Set the horizontal offset of the picture within the cell to 300.
pic.LeftColumnOffset = 300;

// Set the vertical offset of the picture within the cell to 300.
pic.TopRowOffset = 300;
```

---

# spire.xls csharp picture offset
## set picture offset in excel worksheet
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Insert a picture
ExcelPicture pic = sheet.Pictures.Add(2, 2, @"..\..\..\..\..\..\Data\logo.png");

// Set the left offset and top offset of the picture from the current range.
pic.LeftColumnOffset = 200;
pic.TopRowOffset = 100;
```

---

# spire.xls csharp picture reference range
## set reference range for a picture in excel worksheet
```csharp
// Get the first worksheet from the workbook.
Worksheet sheet = workbook.Worksheets[0];

// Set values in cells A1 and B3.
sheet.Range["A1"].Value = "Spire.XLS";
sheet.Range["B3"].Value = "E-iceblue";

// Get the first picture in the worksheet.
ExcelPicture picture = sheet.Pictures[0];

// Set the reference range of the picture to A1:B3.
picture.RefRange = "A1:B3";
```

---

# spire.xls csharp read images
## extract and display images from excel file
```csharp
//Create a Workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.LoadFromFile("ReadImages.xlsx");

//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

//Get the first image
ExcelPicture pic = sheet.Pictures[0];
  
// Show Picture in the PictureBox
using (Form frm1 = new Form())
{
    PictureBox pic1 = new PictureBox();
    pic1.Image = pic.Picture;
    frm1.Width = pic.Picture.Width;
    frm1.Height = pic.Picture.Height;
    frm1.StartPosition = FormStartPosition.CenterParent;
    pic1.Dock = DockStyle.Fill;
    frm1.Controls.Add(pic1);
    frm1.ShowDialog();
}

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp remove picture border
## Remove border from Excel picture
```csharp
// Get the first picture from the first worksheet
ExcelPicture picture = sheet1.Pictures[0];

// Remove the picture border
picture.Line.Visible = false;
```

---

# spire.xls csharp image manipulation
## reset size and position for image in excel worksheet
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Add a picture to the first worksheet.
ExcelPicture picture = sheet.Pictures.Add(1, 1, @"..\..\..\..\..\..\Data\SpireXls.png");

// Set the size for the picture.
picture.Width = 200;
picture.Height = 200;

// Set the position for the picture.
picture.Left = 200;
picture.Top = 100;
```

---

# spire.xls csharp chart image offset
## set image offset for chart background
```csharp
// Add a new worksheet named "Contrast".
Worksheet sheet1 = workbook.Worksheets.Add("Contrast");

// Add chart1 and a background image to sheet1 for comparison.
Chart chart1 = sheet1.Charts.Add(ExcelChartType.ColumnClustered);
chart1.DataRange = sheet.Range["D1:E8"];
chart1.SeriesDataFromRange = false;

// Set the position of the chart.
chart1.LeftColumn = 1;
chart1.TopRow = 11;
chart1.RightColumn = 8;
chart1.BottomRow = 33;

// Add a picture as the background.
chart1.ChartArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Background.png"), "None");
chart1.ChartArea.Fill.Tile = false;

// Set the image offset.
chart1.ChartArea.Fill.PicStretch.Left = 20;
chart1.ChartArea.Fill.PicStretch.Top = 20;
chart1.ChartArea.Fill.PicStretch.Right = 5;
chart1.ChartArea.Fill.PicStretch.Bottom = 5;
```

---

# Spire.XLS C# Image Handling
## Add image to Excel worksheet
```csharp
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Add an image to the specific cell
sheet.Pictures.Add(14, 5, @"..\..\..\..\..\..\Data\SpireXls.png");
```

---

# spire.xls csharp comment
## add comment with author to excel cell
```csharp
// Get the range in which the comment will be added (cell C1).
CellRange range = sheet.Range["C1"];

//Set the author and comment content
string author = "E-iceblue";
string text = "This is demo to show how to add a comment with editable Author property.";

// Add comment to the range and set properties
ExcelComment comment = range.AddComment();
comment.Width = 200;
comment.Visible = true;
comment.Text = string.IsNullOrEmpty(author) ? text : author + ":\n" + text;

// Set the font of the author
ExcelFont font = range.Worksheet.Workbook.CreateFont();
font.FontName = "Tahoma";
font.KnownColor = ExcelColors.Black;
font.IsBold = true;
comment.RichText.SetFont(0, author.Length, font);
```

---

# spire.xls csharp comment
## add comment with picture to excel cell
```csharp
// Set value for the range
sheet.Range["C6"].Text = "E-iceblue";

// Add the comment
ExcelComment comment = sheet.Range["C6"].AddComment();

// Load the image file
Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");

// Fill the comment with a customized background picture
comment.Fill.CustomPicture(image, "logo.png");

// Set the height and width of comment
comment.Height = image.Height;
comment.Width = image.Width;
comment.Visible = true;
```

---

# spire.xls csharp comment
## edit Excel comment
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Load the file from disk.
workbook.LoadFromFile("Template_Xls_8.xlsx");

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Get the first comment.
ExcelComment comment = sheet.Comments[0];

// Edit the comment.
comment.Text = "This comment has been edited by Spire.XLS.";
```

---

# spire.xls csharp name manager
## get comments from name manager in excel
```csharp
// Access the NameRanges property of the workbook
INameRanges nameManager = workbook.NameRanges;

// Create a StringBuilder to store the result
StringBuilder stringBuilder = new StringBuilder();

// Iterate through each name in the NameRanges collection
for (int i = 0; i < nameManager.Count; i++)
{
    // Get the XlsName object at index i
    XlsName name = (XlsName)nameManager[i];

    // Append the name and comment value to the StringBuilder
    stringBuilder.Append("Name: " + name.Name + ", Comment: " + name.CommentValue + "\r\n");
}
```

---

# spire.xls csharp comment visibility
## hide or show excel comments
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Hide the second comment
sheet.Comments[1].IsVisible = false;

// Show the third comment
sheet.Comments[2].IsVisible = true;
```

---

# spire.xls csharp comment
## read excel cell comments
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Show comment in the TextBox
textBox1.Text = sheet.Range["A1"].Comment.Text;
richTextBox1.Rtf = sheet.Range["A2"].Comment.RichText.RtfText;
```

---

# spire.xls csharp excel comments
## remove and modify excel comments
```csharp
//Get all comments of the first sheet
CommentsCollection comments = workbook.Worksheets[0].Comments;

//Change the content of the first comment
comments[0].Text = "This comment has been changed.";

//Remove the second comment
comments[1].Remove();
```

---

# Spire.XLS C# Comment Formatting
## Set fill color for Excel cell comments
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Create Excel font
ExcelFont font = workbook.CreateFont();
font.FontName = "Arial";
font.Size = 11;
font.KnownColor = ExcelColors.Orange;

// Add the comment
CellRange range = sheet.Range["A1"];
range.Comment.Text = "This is a comment";
range.Comment.RichText.SetFont(0, (range.Comment.Text.Length - 1), font);

// Set comment Color
range.Comment.Fill.FillType = ShapeFillType.SolidColor;
range.Comment.Fill.ForeColor = Color.SkyBlue;

// Set comment is visible
range.Comment.Visible = true;
```

---

# spire.xls csharp comment text rotation
## set text rotation for excel cell comments
```csharp
//Create Excel font
ExcelFont font = workbook.CreateFont();
font.FontName = "Arial";
font.Size = 11;
font.KnownColor = ExcelColors.Orange;

//Add the comment
CellRange range = sheet.Range["E1"];
range.Comment.Text = "This is a comment";
range.Comment.RichText.SetFont(0, (range.Comment.Text.Length - 1), font);

// Set its vertical and horizontal alignment 
range.Comment.VAlignment = CommentVAlignType.Center;
range.Comment.HAlignment = CommentHAlignType.Right;

//Set the comment text rotation
range.Comment.TextRotation = TextRotationType.LeftToRight;
```

---

# spire.xls csharp comment position alignment
## Set position and alignment of Excel comments
```csharp
// Create two font styles which will be used in comments
ExcelFont font1 = workbook.CreateFont();
font1.FontName = "Calibri";
font1.Color = Color.Firebrick;
font1.IsBold = true;
font1.Size = 12;
ExcelFont font2 = workbook.CreateFont();
font2.FontName = "Calibri";
font2.Color = Color.Blue;
font2.Size = 12;
font2.IsBold = true;

// Add comment 1 and set its size, text, position and alignment
sheet.Range["G5"].Text = "Spire.XLS";
ExcelComment Comment1 = sheet.Range["G5"].Comment;
Comment1.IsVisible = true;
Comment1.Height = 150;
Comment1.Width = 300;
Comment1.RichText.Text = "Spire.XLS for .Net:\nStandalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. ";
Comment1.RichText.SetFont(0, 19, font1);
Comment1.TextRotation = TextRotationType.LeftToRight;

// Set the position of Comment
Comment1.Top = 20;
Comment1.Left = 40;

// Set the alignment of text in Comment
Comment1.VAlignment = CommentVAlignType.Center;
Comment1.HAlignment = CommentHAlignType.Justified;

// Add comment2 and set its size, text, position and alignment for comparison
sheet.Range["D14"].Text = "E-iceblue";
ExcelComment Comment2 = sheet.Range["D14"].Comment;
Comment2.IsVisible = true;
Comment2.Height = 150;
Comment2.Width = 300;
Comment2.RichText.Text = "About E-iceblue: \nWe focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents.";
Comment2.TextRotation = TextRotationType.LeftToRight;
Comment2.RichText.SetFont(0, 16, font2);

// Set the position of Comment
Comment2.Top = 170;
Comment2.Left = 450;

// Set the alignment of text in Comment
Comment2.VAlignment = CommentVAlignType.Top;
Comment2.HAlignment = CommentHAlignType.Justified;
```

---

# spire.xls csharp comment
## write regular and rich text comments to Excel cells
```csharp
// Creates fonts for styling comments
ExcelFont font = workbook.CreateFont();
font.FontName = "Arial";
font.Size = 11;
font.KnownColor = ExcelColors.Orange;
ExcelFont fontBlue = workbook.CreateFont();
fontBlue.KnownColor = ExcelColors.LightBlue;
ExcelFont fontGreen = workbook.CreateFont();
fontGreen.KnownColor = ExcelColors.LightGreen;

// Get cell B11 and add a regular comment
CellRange range = sheet.Range["B11"];
range.Text = "Regular comment";
range.Comment.Text = "Regular comment";
range.AutoFitColumns();

// Get cell B12 and add a rich text comment
range = sheet.Range["B12"];
range.Text = "Rich text comment";
range.RichText.SetFont(0, 16, font);
range.AutoFitColumns();

// Set rich text comment with different colored fonts
range.Comment.RichText.Text = "Rich text comment";
range.Comment.RichText.SetFont(0, 4, fontGreen);
range.Comment.RichText.SetFont(5, 9, fontBlue);
```

---

# Spire.XLS C# Chart Sheet to SVG Conversion
## Convert Excel chart sheet to SVG format

```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the document
workbook.LoadFromFile(inputFile);

//Get the chartsheet by name
ChartSheet cs = workbook.GetChartSheetByName("Chart1");

//Convert chart sheet to SVG
FileStream fs = new FileStream(outputFile, FileMode.Create);
cs.ToSVGStream(fs);
fs.Flush();
fs.Close();

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp csv to datatable
## convert CSV file to DataTable using Spire.XLS
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.LoadFromFile("CSVSample.csv", ",");

//Get the first worksheet
Worksheet worksheet = workbook.Worksheets[0];

//Export to datatable
System.Data.DataTable dataTable = worksheet.ExportDataTable();

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csv to excel conversion
## convert CSV file to Excel format
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load a csv file
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CSVToExcel.csv", ",", 1, 1);

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Ignore error options for the range D2:E19, treating numbers as text
sheet.Range["D2:E19"].IgnoreErrorOptions = IgnoreErrorType.NumberAsText;

// Auto-fit columns in the allocated range of the worksheet
sheet.AllocatedRange.AutoFitColumns();

//Save the file 
workbook.SaveToFile("CSVToExcel_result.xlsx", ExcelVersion.Version2013);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp csv to pdf conversion
## convert CSV file to PDF format using Spire.XLS
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.LoadFromFile("CSVSample.csv", ",", 1, 1);

//Set the SheetFitToPage property as true
workbook.ConverterSetting.SheetFitToPage = true;

//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Autofit a column if the characters in the column exceed column width
for (int i = 1; i < sheet.Columns.Length; i++)
{
    sheet.AutoFitColumn(i);
}

//Save to PDF document
workbook.SaveToFile("CSVToPDF.pdf", FileFormat.PDF);
```

---

# Spire.XLS C# Worksheet to PDF Conversion
## Convert each worksheet in an Excel workbook to a separate PDF file
```csharp
//Save each sheet to PDF
foreach (Worksheet sheet in workbook.Worksheets)
{
    string FileName = sheet.Name + ".pdf";
    //Save the sheet to PDF
    sheet.SaveToPdf(FileName);
}
```

---

# spire.xls csharp embed non-installed fonts
## embed custom fonts in excel chart elements
```csharp
// Load the font file from disk
workbook.CustomFontFilePaths = new string[] { @"..\..\..\..\..\..\Data\PT_Serif-Caption-Web-Regular.ttf" };
System.Collections.Hashtable result = workbook.GetCustomFontParsedResult(); 

ArrayList valueList = new ArrayList(result.Values);

// Apply the font for PrimaryValueAxis of chart
chart.PrimaryValueAxis.Font.FontName = valueList[0] as string;

// Apply the font for PrimaryCategoryAxis of chart
chart.PrimaryCategoryAxis.Font.FontName = valueList[0] as string;

// Apply the font for the first chartSerie of chart
ChartSerie chartSerie1 = chart.Series[0];
chartSerie1.DataPoints.DefaultDataPoint.DataLabels.FontName = valueList[0] as string;
```

---

# Spire.XLS C# Excel to Markdown Conversion
## Convert Excel files to Markdown format using Spire.XLS library
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelToMarkdown.xlsx");

// Save to Markdown document
string output = "ExcelToMarkdown_out.md";
workbook.SaveToFile(output, FileFormat.Markdown);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp conversion
## fit width when converting excel to pdf
```csharp
// Create a Workbook
Workbook workbook = new Workbook();

// Process each worksheet
foreach (Worksheet sheet in workbook.Worksheets)
{
    // Auto fit page height
    sheet.PageSetup.FitToPagesTall = 0;
    // Fit one page width
    sheet.PageSetup.FitToPagesWide = 1;
}

// Convert to PDF
workbook.SaveToFile(result, FileFormat.PDF);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS HTML to Excel Conversion
## Convert HTML content to Excel format using Spire.XLS library
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load html
workbook.LoadFromHtml(htmlFilePath);

//Save to Excel file
workbook.SaveToFile(outputFilePath, ExcelVersion.Version2013);
```

---

# spire.xls csharp file conversion
## load and save .et and .ett files (Kingsoft Spreadsheets formats)
```csharp
//create a workbook
Workbook workbook = new Workbook();

//load .et or .ett file 
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Sample-et.et");
//workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Sample-ett.ett");

//save to .et or .ett file
workbook.SaveToFile("result.et", FileFormat.ET);
//workbook.SaveToFile("result.ett", FileFormat.ETT);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS XML to Excel Conversion
## Convert Office Open XML format to Excel file format
```csharp
// Create a workbook 
Workbook workbook = new Workbook();

// Load from XML
workbook.LoadFromXml(fileStream);

// Save to Excel
workbook.SaveToFile(outputFileName, excelVersion);

// Dispose resources
workbook.Dispose();
```

---

# Spire.XLS C# Excel to PDF Conversion
## Convert a selected range from Excel to PDF
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConversionSample1.xlsx");

//Add a new sheet to workbook
workbook.Worksheets.Add("newsheet");

//Copy your area to new sheet.
workbook.Worksheets[0].Range["A9:E15"].Copy(workbook.Worksheets[1].Range["A9:E15"], false, true);

//Auto fit column width
workbook.Worksheets[1].Range["A9:E15"].AutoFitColumns();

//Save the document
string output = "SelectedRangeToPDF.pdf";
workbook.Worksheets[1].SaveToPdf(output);
```

---

# Spire.XLS C# Global Custom Fonts
## Demonstrates how to set global custom fonts in Spire.XLS
```csharp
// Set custom font directory
string[] fontPath = { @"..\..\..\..\..\..\Data\fonts" };

// Create a new workbook object
Workbook workbook = new Workbook();
Workbook.SetGlobalCustomFontsFolders(fontPath);
```

---

# spire.xls csharp sheet to emf conversion
## convert excel worksheet to emf image format
```csharp
//Create a memory stream
MemoryStream stream = new MemoryStream();

//Save excel worksheet into EMF stream
sheet.ToEMFStream(stream, 1, 1, 28, 8, EmfType.EmfPlusDual);

//Create image from the stream
Image img = Image.FromStream(stream);
```

---

# spire.xls csharp sheet to image
## Convert Excel worksheet to image
```csharp
//Get the first worksheet in excel workbook
Worksheet sheet = workbook.Worksheets[0];

// Save to image
sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn).Save("SheetToImage.png");
```

---

# Spire.XLS C# Excel to Image Conversion
## Convert specific cell ranges to images in different formats
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConversionSample1.xlsx");

//Get the first worksheet in Excel file
Worksheet sheet = workbook.Worksheets[0];

//Specify Cell Ranges and Save to certain Image formats
sheet.ToImage(1, 1, 7, 5).Save("image1.png", ImageFormat.Png);
sheet.ToImage(8, 1, 15, 5).Save("image2.jpg", ImageFormat.Jpeg);
sheet.ToImage(17, 1, 23, 5).Save("image3.bmp", ImageFormat.Bmp);
```

---

# Spire.XLS C# Font Directory Specification
## Set custom font directory for Excel to PDF conversion
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Specify font directory
workbook.CustomFontFileDirectory = new string[] { (@"..\..\..\..\..\..\Data\Font") };
```

---

# spire.xls csharp excel to csv conversion
## convert Excel worksheet to CSV format
```csharp
//get the first sheet
Worksheet sheet = workbook.Worksheets[0];

//convert to CSV file
sheet.SaveToFile("ToCSV.csv", ",", Encoding.UTF8);
```

---

# Excel to CSV conversion with double quotes
## Demonstrates how to convert Excel files to CSV format with double quotes using Spire.XLS
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Convert to CSV file,
// When the last parameter is set to true, there are double quotes. The default parameter is false
workbook.SaveToFile("output.csv", ",", true);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp excel to csv conversion
## Convert Excel worksheet to CSV with filtered values
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load excel document from disk
workbook.LoadFromFile("AutofilterSample.xlsx");

// Convert to CSV file with filtered value
workbook.Worksheets[0].SaveToFile("ToCSVWithFilteredValue.csv", ";", false);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp excel to encrypted pdf
## Convert Excel file to encrypted PDF format
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Load a excel document
workbook.LoadFromFile("input.xlsx");

// Set open and permission password to encrypt converted pdf
workbook.ConverterSetting.PdfSecurity.Encrypt("123","456", PdfPermissionsFlags.Print, PdfEncryptionKeySize.Key128Bit);

// Convert excel to pdf
workbook.SaveToFile("output.pdf", FileFormat.PDF);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls excel to html conversion
## Convert Excel worksheet to HTML format using Spire.XLS
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load file from disk
workbook.LoadFromFile("ToHtml.xlsx");

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Create HTML options for saving to HTML format
HTMLOptions options = new HTMLOptions();

// Embed images in the HTML file
options.ImageEmbedded = true;

// Save the file 
sheet.SaveToHtml("sample.html", options);
```

---

# spire.xls csharp excel to html
## convert excel worksheet to html stream
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the Excel document from disk
workbook.LoadFromFile("ReadImages.xlsx");

//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Set the html options
HTMLOptions options = new HTMLOptions();
options.ImageEmbedded = true;

//Save sheet to html stream
FileStream fileStream = new FileStream("Output.html", FileMode.Create);
sheet.SaveToHtml(fileStream, options);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS C# HTML Conversion
## Convert Excel to HTML with hidden worksheets
```csharp
// Create a workbook
Workbook book = new Workbook();

// Load the document
book.LoadFromFile("ToHtmlWithHiddenWorksheets.xlsx");

// Save Excel to Html
// false --- To Html with the hidden Worksheet
// true--- To Html without the hidden Worksheet
string result = "result.html";
book.SaveToHtml(result, false);

// Dispose of the workbook object to release resources
book.Dispose();
```

---

# spire.xls csharp conversion
## convert excel worksheet to high resolution image
```csharp
//Convert the worksheet to EMF stream
using (MemoryStream ms = new MemoryStream())
{
    worksheet.ToEMFStream(ms, 1, 1, worksheet.LastRow, worksheet.LastColumn);

    //Create an image from the EMF stream
    Image image = Image.FromStream(ms);
    Bitmap images = ResetResolution(image as Metafile, 300);
}

//A custom function to reset the image resolution
private static Bitmap ResetResolution(Metafile mf, float resolution)
{
    int width = (int)(mf.Width * resolution / mf.HorizontalResolution);
    int height = (int)(mf.Height * resolution / mf.VerticalResolution);
    Bitmap bmp = new Bitmap(width, height);
    bmp.SetResolution(resolution, resolution);
    Graphics g = Graphics.FromImage(bmp);
    g.DrawImage(mf, 0, 0);
    g.Dispose();
    return bmp;
}
```

---

# spire.xls csharp excel to image
## convert excel worksheet to image without white space
```csharp
// Set the margin as 0 to remove the white space around the image
sheet.PageSetup.LeftMargin = 0;
sheet.PageSetup.BottomMargin = 0;
sheet.PageSetup.TopMargin = 0;
sheet.PageSetup.RightMargin = 0;

// Convert to image
Image image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn);
```

---

# Spire.XLS C# Excel to ODS Conversion
## Convert Excel files to ODS format using Spire.XLS library
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load an Excel document
workbook.LoadFromFile("ToODS.xlsx");

// Convert to ODS file
workbook.SaveToFile("Result.ods", FileFormat.ODS);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp excel to ofd conversion
## convert excel file to ofd format
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToOFD.xlsx");

// Save to ofd file
workbook.SaveToFile("result.ofd", FileFormat.OFD);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp conversion
## convert Excel workbook to Office Open XML format
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Set the text "Hello World" in cell A1 of the worksheet
sheet.Range["A1"].Text = "Hello World";

// Apply the color Gray25Percent to cell B1 using a known color
sheet.Range["B1"].Style.KnownColor = ExcelColors.Gray25Percent;

// Apply the color Gold to cell C1 using a known color
sheet.Range["C1"].Style.KnownColor = ExcelColors.Gold;

// Save the workbook as an XML file
workbook.SaveAsXml("sample.xml");
```

---

# Spire.XLS C# Excel to PDF Conversion
## Convert Excel files to PDF format using Spire.XLS library
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load Excel file
workbook.LoadFromFile("input.xlsx");

// Set the ConverterSetting property to enable fitting sheets to page during PDF conversion
workbook.ConverterSetting.SheetFitToPage = true;

// Save the workbook as a PDF file
workbook.SaveToFile("output.pdf", FileFormat.PDF);
```

---

# Spire.XLS C# PDF/A-1B Conversion
## Convert Excel to PDF/A-1B format
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Convert excel to PDFA/1-B
workbook.ConverterSetting.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;

// Save the document as PDF
workbook.SaveToFile("ToPDFA1B_result.pdf", FileFormat.PDF);
```

---

# spire.xls csharp excel to pdf
## simple excel to pdf conversion
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load a excel document
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF.xlsx");

// Convert excel to pdf
string result = "ToPdfSimply_result.pdf";
workbook.SaveToFile(result, FileFormat.PDF);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS CSharp Convert to PDF with Page Size
## Convert Excel to PDF with custom page size
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

foreach (Worksheet sheet in workbook.Worksheets)
{
    // Change the page size
    sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
}

// Convert workbook to PDF
workbook.SaveToFile("result.pdf", FileFormat.PDF);
```

---

# spire.xls csharp pdf conversion
## convert excel to pdf with custom page size
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Load the document from disk
workbook.LoadFromFile("SampleB_2.xlsx");

foreach (Worksheet sheet in workbook.Worksheets)
{
    // Change the page size
    sheet.PageSetup.SetCustomPaperSize(100f, 100f);
}

// Save the result file
string result = "result.pdf";
workbook.SaveToFile(result, FileFormat.PDF);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp conversion
## convert Excel to PostScript format
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load file from disk
workbook.LoadFromFile("ToPostScript.xlsx");

// Convert to PostScript file
string result = "Result.ps";
workbook.SaveToFile(result, FileFormat.PostScript);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp excel to html conversion
## convert excel file to standalone html
```csharp
// Create a new Workbook object
Workbook workbook = new Workbook();

// Set the HTMLOptions to create a standalone HTML file
HTMLOptions.Default.IsStandAloneHtmlFile = true;

// Save the Excel document as an HTML file
workbook.SaveToStream(fileStream, FileFormat.HTML);
```

---

# spire.xls csharp excel to svg conversion
## convert excel worksheets to svg format
```csharp
// Iterate through each worksheet in the workbook
for (int i = 0; i < workbook.Worksheets.Count; i++)
{
    // Create a FileStream to write the SVG content to a file
    FileStream fs = new FileStream(string.Format("sheet{0}.svg", i), FileMode.Create);
    // Convert the worksheet to SVG and write it to the FileStream
    workbook.Worksheets[i].ToSVGStream(fs, 0, 0, 0, 0);
    // Flush and close the FileStream to ensure data is written and resources are released
    fs.Flush();
    fs.Close();
}
```

---

# spire.xls csharp excel to text conversion
## convert excel worksheet to text file
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet in excel workbook
Worksheet sheet = workbook.Worksheets[0];

//Save to text
string output = "ExceltoTxt.txt";
sheet.SaveToFile(output, " ", Encoding.UTF8);
```

---

# Spire.XLS C# Excel to TIFF Conversion
## Convert Excel workbook to multi-page TIFF image
```csharp
// Convert workbook worksheets to image array
private static Image[] ToImage(Workbook workbook)
{
    //Get the worksheet count of workbook
    int workSheetNo = workbook.Worksheets.Count;

    //Create an array
    Image[] images = new Image[workSheetNo];

    //Save worksheet to image and add the array
    for (int i = 0; i < workSheetNo; i++)
    {
        Worksheet workSheet = workbook.Worksheets[i];
        string output = string.Format("result{0}.jpg",i+1);
        workSheet.SaveToImage(output);
        Image image= Image.FromFile(output);
        images[i] = image;
    }
    return images;
}

// Get image encoder information for specified MIME type
private static ImageCodecInfo GetEncoderInfo(string mimeType)
{
    ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();
    for (int j = 0; j < encoders.Length; j++)
    {
        if (encoders[j].MimeType == mimeType)
            return encoders[j];
    }
    throw new Exception(mimeType + " mime type not found in ImageCodecInfo");
}

// Join multiple images into a single TIFF file
public static void JoinTiffImages(Image[] images, string outFile, EncoderValue compressEncoder)
{
    //Use the save encoder
    Encoder enc = Encoder.SaveFlag;
    EncoderParameters ep = new EncoderParameters(2);
    ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);
    ep.Param[1] = new EncoderParameter(Encoder.Compression, (long)compressEncoder);
    Image pages = images[0];
    int frame = 0;
    ImageCodecInfo info = GetEncoderInfo("image/tiff");
    foreach (Image img in images)
    {
        if (frame == 0)
        {
            pages = img;
            //save the first frame
            pages.Save(outFile, info, ep);
        }

        else
        {
            //save the intermediate frames
            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);

            pages.SaveAdd(img, ep);
        }
        if (frame == images.Length - 1)
        {
            //flush and close.
            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.Flush);
            pages.SaveAdd(ep);
        }
        frame++;
    }
}
```

---

# Spire.XLS C# Excel to UOS Conversion
## Convert Excel file to UOS format using Spire.XLS library
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the Excel document
workbook.LoadFromFile(inputFilePath);

// Save to UOS format
workbook.SaveToFile(outputFilePath, FileFormat.UOS);
```

---

# Spire.XLS C# Excel to XPS Conversion
## Convert Excel file to XPS format using Spire.XLS library
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load a file from the specified path into the workbook
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToXPS.xlsx");

// Save the workbook as an XPS file with the name "ToXPS.xps" using the Spire.Xls library's XPS file format
workbook.SaveToFile("ToXPS.xps", Spire.Xls.FileFormat.XPS);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls uos to excel conversion
## convert uos file to excel format
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.LoadFromFile("input.uos", ExcelVersion.UOS);

//Save to excel file
workbook.SaveToFile("output.xlsx", ExcelVersion.Version2013);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp workbook to html conversion
## Convert Excel workbook to HTML format
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Convert to html
workbook.SaveToHtml("result.html");

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp xlsb conversion
## convert and manipulate XLSB files using Spire.XLS
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load a file from the specified path into the workbook
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\XLSB.xlsb");

// Get the first worksheeet
Worksheet sheet = workbook.Worksheets[0];

// Export data to data table
this.dataGrid1.DataSource = sheet.ExportDataTable();

// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Insert data from a data table into the worksheet, starting from cell A1
sheet.InsertDataTable((DataTable)this.dataGrid1.DataSource, true, 1, 1, -1, -1);

// Define cell styles for odd and even rows
CellStyle oddStyle = workbook.Styles.Add("oddStyle");
oddStyle.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
oddStyle.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
oddStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
oddStyle.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
oddStyle.KnownColor = ExcelColors.LightGreen1;

CellStyle evenStyle = workbook.Styles.Add("evenStyle");
evenStyle.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
evenStyle.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
evenStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
evenStyle.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
evenStyle.KnownColor = ExcelColors.LightTurquoise;

// Apply the odd and even styles to the rows in the allocated range of the worksheet
foreach (CellRange range in sheet.AllocatedRange.Rows)
{
    if (range.Row % 2 == 0)
        range.CellStyleName = evenStyle.Name;
    else
        range.CellStyleName = oddStyle.Name;
}

// Set the header row style
CellStyle styleHeader = sheet.AllocatedRange.Rows[0].Style;
styleHeader.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
styleHeader.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
styleHeader.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
styleHeader.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
styleHeader.VerticalAlignment = VerticalAlignType.Center;
styleHeader.KnownColor = ExcelColors.Green;
styleHeader.Font.KnownColor = ExcelColors.White;
styleHeader.Font.IsBold = true;

// Autofit columns and rows in the allocated range of the worksheet
sheet.AllocatedRange.AutoFitColumns();
sheet.AllocatedRange.AutoFitRows();

// Set the height of the first row to 20
sheet.Rows[0].RowHeight = 20;

// Save the workbook as an XLSB file
workbook.SaveToFile("sample.xlsb", ExcelVersion.Xlsb2010);
```

---

# spire.xls csharp file conversion
## convert XLS to XLSM format
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MacroSample.xls",ExcelVersion.Version97to2003);

//Save the workbook as a new XLSM file
string output = "XLSToXLSM.xlsm";
workbook.SaveToFile(output);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp autofilter
## apply autofilter to blank cells in excel worksheet
```csharp
// Match the blank data
sheet.AutoFilters.MatchBlanks(0);

// Filter
sheet.AutoFilters.Filter();
```

---

# Spire.XLS C# AutoFilter
## Apply auto-filter to non-blank cells in Excel
```csharp
//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Match the non blank data
sheet.AutoFilters.MatchNonBlanks(0);

//Filter
sheet.AutoFilters.Filter();
```

---

# spire.xls csharp excel filter
## create auto filter in excel worksheet
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Specify the range for creating the filter (in this case, A1 to J1)
sheet.AutoFilters.Range = sheet.Range["A1:J1"];
```

---

# spire.xls csharp data validation
## create data validation for excel cells
```csharp
// Decimal DataValidation
sheet.Range["B11"].Text = "Input Number(3-6):";
CellRange rangeNumber = sheet.Range["B12"];
rangeNumber.DataValidation.CompareOperator = ValidationComparisonOperator.Between;
rangeNumber.DataValidation.Formula1 = "3";
rangeNumber.DataValidation.Formula2 = "6";
rangeNumber.DataValidation.AllowType = CellDataType.Decimal;
rangeNumber.DataValidation.ErrorMessage = "Please input correct number!";
rangeNumber.DataValidation.ShowError = true;
rangeNumber.Style.KnownColor = ExcelColors.Gray25Percent;

// Date DataValidation
sheet.Range["B14"].Text = "Input Date:";
CellRange rangeDate = sheet.Range["B15"];
rangeDate.DataValidation.AllowType = CellDataType.Date;
rangeDate.DataValidation.CompareOperator = ValidationComparisonOperator.Between;
rangeDate.DataValidation.Formula1 = "1/1/1970";
rangeDate.DataValidation.Formula2 = "12/31/1970";
rangeDate.DataValidation.ErrorMessage = "Please input correct date!";
rangeDate.DataValidation.ShowError = true;
rangeDate.DataValidation.AlertStyle = AlertStyleType.Warning;
rangeDate.Style.KnownColor = ExcelColors.Gray25Percent;

// TextLength DataValidation
sheet.Range["B17"].Text = "Input Text:";
CellRange rangeTextLength = sheet.Range["B18"];
rangeTextLength.DataValidation.AllowType = CellDataType.TextLength;
rangeTextLength.DataValidation.CompareOperator = ValidationComparisonOperator.LessOrEqual;
rangeTextLength.DataValidation.Formula1 = "5";
rangeTextLength.DataValidation.ErrorMessage = "Enter a Valid String!";
rangeTextLength.DataValidation.ShowError = true;
rangeTextLength.DataValidation.AlertStyle = AlertStyleType.Stop;
rangeTextLength.Style.KnownColor = ExcelColors.Gray25Percent;

// Auto fit the column width for better visibility
sheet.AutoFitColumn(2);
```

---

# Spire.XLS C# Filter by String
## Filter Excel cells by string pattern
```csharp
// Set the range for filtering cells data, in this case, column D from row 1 to 19
sheet.AutoFilters.Range = sheet.Range["D1:D19"];

// Get the filter column for custom filtering
FilterColumn filtercolumn = (FilterColumn)sheet.AutoFilters[0];

// Apply a custom filter to display only cells starting with "South"
sheet.AutoFilters.CustomFilter(filtercolumn, FilterOperatorType.Equal, "South*");

// Apply the filters
sheet.AutoFilters.Filter();
```

---

# Spire.XLS C# Data Validation
## Get settings of data validation from Excel cell
```csharp
//Get first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];

//Cell B4 has the Decimal Validation
CellRange cell = worksheet.Range["B4"];

//Get the valditation of this cell
Validation validation = cell.DataValidation;

//Get the settings
string allowType = validation.AllowType.ToString();
string data = validation.CompareOperator.ToString();
string minimum = validation.Formula1.ToString();
string maximum = validation.Formula2.ToString();
string ignoreBlank = validation.IgnoreBlank.ToString();

//Set string format for displaying
string result = string.Format("Settings of Validation: \r\nAllow Type: " + allowType + "\r\nData: " + data + "\r\nMinimum: " + minimum +"\r\nMaximum: " + maximum + "\r\nIgnoreBlank: "+ignoreBlank);
```

---

# spire.xls csharp list data validation
## create list data validation in excel cell
```csharp
//Set data validation for cell
CellRange range = sheet.Range["D10"];
range.DataValidation.ShowError = true;
range.DataValidation.AlertStyle = AlertStyleType.Stop;
range.DataValidation.ErrorTitle = "Error";
range.DataValidation.ErrorMessage = "Please select a city from the list";
range.DataValidation.DataRange = sheet.Range["A7:A10"];
```

---

# Remove Auto Filters from Excel Worksheet
## This code demonstrates how to remove auto filters from an Excel worksheet using Spire.XLS.
```csharp
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Remove the auto filters.
sheet.AutoFilters.Clear();
```

---

# spire.xls csharp data validation
## remove data validation from excel ranges
```csharp
//Create an array of rectangles, which is used to locate the ranges in worksheet.
Rectangle[] rectangles = new Rectangle[1];

//Assign value to the first element of the array. This rectangle specifies the cells from A1 to B3.
rectangles[0] = new Rectangle(0, 0, 1, 2);

//Remove validations in the ranges represented by rectangles.
workbook.Worksheets[0].DVTable.Remove(rectangles);
```

---

# Spire.XLS C# Data Validation
## Set data validation on separate sheet
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Access the first sheet in the workbook
Worksheet sheet1 = workbook.Worksheets[0];

// Access the second sheet in the workbook
Worksheet sheet2 = workbook.Worksheets[1];

// Enable the option to allow data from a different sheet in data validation
sheet2.ParentWorkbook.Allow3DRangesInDataValidation = true;

// Set the data range for data validation on cell B11 of the first sheet,
// using the range A1:A7 from the second sheet as the source of data
sheet1.Range["B11"].DataValidation.DataRange = sheet2.Range["A1:A7"];
```

---

# spire.xls csharp time data validation
## set time data validation for excel cells
```csharp
//Set Time data validation for cell "D12"
CellRange range = sheet.Range["D12"];
range.DataValidation.AllowType = CellDataType.Time;
range.DataValidation.CompareOperator = ValidationComparisonOperator.Between;
range.DataValidation.Formula1 = "09:00";
range.DataValidation.Formula2 = "18:00";
range.DataValidation.AlertStyle = AlertStyleType.Info;
range.DataValidation.ShowError = true;
range.DataValidation.ErrorTitle = "Time Error";
range.DataValidation.ErrorMessage = "Please enter a valid time";
range.DataValidation.InputMessage = "Time Validation Type";
range.DataValidation.IgnoreBlank = true;
range.DataValidation.ShowInput = true;
```

---

# spire.xls data validation
## verify data against cell validation criteria
```csharp
//Get first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];

//Cell B4 has the Decimal Validation
CellRange cell = worksheet.Range["B4"];

//Get the validation of this cell
Validation validation = cell.DataValidation;

//Get the specified data range
double minimum = double.Parse(validation.Formula1);
double maximum = double.Parse(validation.Formula2);

//Set different numbers for the cell
for (int i = 5; i < 100; i=i+40 )
{
    cell.NumberValue = i;
    string result=null;
    //Verify 
    if (cell.NumberValue < minimum || cell.NumberValue > maximum)
    {
        //Set string format for displaying
        result = string.Format("Is input "+ i +" a valid value for this Cell: false");
    }
    else
    {
        //Set string format for displaying
        result = string.Format("Is input " + i + " a valid value for this Cell: true");
    }
    //Add result string to StringBuilder
    content.AppendLine(result);
}
```

---

# Spire.XLS C# Data Validation
## Implement whole number data validation in Excel cells
```csharp
// Set the text in cell C12 for the data validation prompt
sheet.Range["C12"].Text = "Please enter a number between 10 and 100:";

// Auto-fit the columns to adjust the width
sheet.Range["C12"].AutoFitColumns();

// Set Whole Number data validation for cell D12
CellRange range = sheet.Range["D12"];
range.DataValidation.AllowType = CellDataType.Integer;
range.DataValidation.CompareOperator = ValidationComparisonOperator.Between;
range.DataValidation.Formula1 = "10";
range.DataValidation.Formula2 = "100";
range.DataValidation.AlertStyle = AlertStyleType.Info;
range.DataValidation.ShowError = true;
range.DataValidation.ErrorTitle = "Error";
range.DataValidation.ErrorMessage = "Please enter a valid number";
range.DataValidation.InputMessage = "Whole Number Validation Type";
range.DataValidation.IgnoreBlank = true;
range.DataValidation.ShowInput = true;
```

---

# spire.xls csharp chart
## Add data table to chart
```csharp
//Get the first chart
Chart chart = sheet.Charts[0];

// Enable the data table for the chart
chart.HasDataTable = true;
```

---

# spire.xls csharp chart
## add picture to chart
```csharp
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Get the first chart
Chart chart = sheet.Charts[0];

// Add the picture in chart
chart.Shapes.AddPicture(@"..\..\..\..\..\..\Data\SpireXls.png");
```

---

# spire.xls csharp chart textbox
## add textbox to chart in excel
```csharp
// Get the first chart
Chart chart = sheet.Charts[0];

// Add a Textbox
ITextBoxLinkShape textbox = chart.Shapes.AddTextBox();

// Set the width of the textbox
textbox.Width = 1200;
// Set the height of the textbox
textbox.Height = 320;
// Set the left position of the textbox
textbox.Left = 1000;
// Set the top position of the textbox
textbox.Top = 480;
textbox.Text = "This is a textbox";
```

---

# spire.xls csharp trendline
## add different types of trendlines to excel charts
```csharp
//select chart and set logarithmic trendline
Chart chart = sheet.Charts[0];
chart.ChartTitle = "Logarithmic Trendline";
chart.Series[0].TrendLines.Add(TrendLineType.Logarithmic);

//select chart and set moving_average trendline
Chart chart1 = sheet.Charts[1];
chart1.ChartTitle = "Moving Average Trendline";
chart1.Series[0].TrendLines.Add(TrendLineType.Moving_Average);

//select chart and set linear trendline
Chart chart2 = sheet.Charts[2];
chart2.ChartTitle = "Linear Trendline";
chart2.Series[0].TrendLines.Add(TrendLineType.Linear);

//select chart and set exponential trendline
Chart chart3 = sheet.Charts[3];
chart3.ChartTitle = "Exponential Trendline";
chart3.Series[0].TrendLines.Add(TrendLineType.Exponential);
```

---

# spire.xls csharp chart
## adjust bar space in chart
```csharp
//Get the first worksheet from workbook and then get the first chart from the worksheet
Worksheet ws = workbook.Worksheets[0];
Chart chart = ws.Charts[0];

//Adjust the space between bars
foreach (ChartSerie cs in chart.Series)
{
    cs.Format.Options.GapWidth = 200;
    cs.Format.Options.Overlap = 0;
}
```

---

# spire.xls csharp chart effect
## apply soft edges effect to chart
```csharp
//Get the chart
Chart chart = sheet.Charts[0];

//Specify the size of the soft edge. Value can be set from 0 to 100
chart.ChartArea.Shadow.SoftEdge = 25;
```

---

# Spire.XLS C# Chart Size and Position
## Change chart size and position in Excel
```csharp
//Get the chart
Chart chart = sheet.Charts[0];

//Change chart size
chart.Width = 600;
chart.Height = 500;

//Change chart position
chart.LeftColumn = 3;
chart.TopRow = 7;
```

---

# spire.xls csharp chart
## change data label in chart
```csharp
//Get the chart
Chart chart = sheet.Charts[0];

//Change data label of the first datapoint of the first series
chart.Series[0].DataPoints[0].DataLabels.Text = "changed data label";
```

---

# spire.xls csharp chart
## Change chart data range
```csharp
// Get the first worksheet 
Worksheet sheet = workbook.Worksheets[0];

// Get chart
Chart chart = sheet.Charts[0];

// Change data range
chart.DataRange = sheet.Range["A1:C4"];
```

---

# spire.xls csharp chart gridlines
## change color of major gridlines in excel chart
```csharp
// Get the chart
Chart chart = sheet.Charts[0];

// Change the color of major gridlines
chart.PrimaryValueAxis.MajorGridLines.LineProperties.Color = Color.Red;
```

---

# spire.xls csharp chart series color
## change chart series color
```csharp
//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Get the first chart
Chart chart = sheet.Charts[0];

//Get the second series
ChartSerie cs = chart.Series[1];

//Set the fill type
cs.Format.Fill.FillType = ShapeFillType.SolidColor;

//Change the fill color
cs.Format.Fill.ForeColor = Color.Orange;
```

---

# spire.xls csharp chart axis title
## Set titles for chart axes in Excel
```csharp
//Get the chart
Chart chart = sheet.Charts[0];

//Set axis title
chart.PrimaryCategoryAxis.Title = "Category Axis";
chart.PrimaryValueAxis.Title = "Value axis";

//Set font size
chart.PrimaryCategoryAxis.Font.Size = 12;
chart.PrimaryValueAxis.Font.Size = 12;
```

---

# spire.xls csharp chart to emf
## Convert Excel chart to EMF image format
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Save chart as Emf image
using (MemoryStream stream = new MemoryStream())
{
    workbook.SaveChartAsEmfImage(workbook.Worksheets[0], 0, stream);
    File.WriteAllBytes("EmfImage.emf", stream.ToArray());
}

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp chart to image
## convert excel chart to image
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load file from disk
workbook.LoadFromFile("ChartToImage.xlsx");

// Save chart as image
Image image= workbook.SaveChartAsImage(workbook.Worksheets[0], 0);
```

---

# spire.xls csharp box and whisker chart
## create Box and Whisker chart with specific settings for each series
```csharp
// Add a new chart
Chart officeChart = sheet.Charts.Add();

//Set the chart title
officeChart.ChartTitle = "Yearly Vehicle Sales";

// Set chart type as Box and Whisker
officeChart.ChartType = ExcelChartType.BoxAndWhisker;

// Set data range in the worksheet
officeChart.DataRange = sheet["A1:E17"];

// Box and Whisker settings on first series
ChartSerie seriesA = officeChart.Series[0];
seriesA.DataFormat.ShowInnerPoints = false;
seriesA.DataFormat.ShowOutlierPoints = true;
seriesA.DataFormat.ShowMeanMarkers = true;
seriesA.DataFormat.ShowMeanLine = false;
seriesA.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian;

// Box and Whisker settings on second series   
ChartSerie seriesB = officeChart.Series[1];
seriesB.DataFormat.ShowInnerPoints = false;
seriesB.DataFormat.ShowOutlierPoints = true;
seriesB.DataFormat.ShowMeanMarkers = true;
seriesB.DataFormat.ShowMeanLine = false;
seriesB.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.InclusiveMedian;

// Box and Whisker settings on third series   
ChartSerie seriesC = officeChart.Series[2];
seriesC.DataFormat.ShowInnerPoints = false;
seriesC.DataFormat.ShowOutlierPoints = true;
seriesC.DataFormat.ShowMeanMarkers = true;
seriesC.DataFormat.ShowMeanLine = false;
seriesC.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian;
```

---

# spire.xls csharp bubble chart
## create bubble chart in excel using spire.xls
```csharp
// Add a Bubble chart to the worksheet
Chart chart = sheet.Charts.Add(ExcelChartType.Bubble);

// Set the title of the chart
chart.ChartTitle = "Bubble";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

// Specify the range of data for the chart
chart.DataRange = sheet.Range["A1:C5"];
chart.SeriesDataFromRange = false;

// Set the range of values for the bubbles in the chart
chart.Series[0].Bubbles = sheet.Range["C2:C5"];

// Set the position of the chart on the worksheet
chart.LeftColumn = 7;
chart.TopRow = 6;
chart.RightColumn = 16;
chart.BottomRow = 29;
```

---

# spire.xls csharp pivot chart
## create chart based on pivot table
```csharp
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets[0];

// Get the pivot table
XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

// Add a chart based on the pivot table to the second worksheet
workbook.Worksheets[1].Charts.Add(ExcelChartType.BarClustered, pt);
```

---

# spire.xls csharp doughnut chart
## create a doughnut chart with percentage labels
```csharp
// Add a new chart and set its type to Doughnut
Chart chart = sheet.Charts.Add();
chart.ChartType = ExcelChartType.Doughnut;

// Set the data range for the chart
chart.DataRange = sheet.Range["A1:B5"];
chart.SeriesDataFromRange = false;

// Set the position of the chart on the worksheet
chart.LeftColumn = 4;
chart.TopRow = 2;
chart.RightColumn = 12;
chart.BottomRow = 22;

// Set the chart title
chart.ChartTitle = "Market share by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

// Enable percentage labels for each data point
foreach (ChartSerie cs in chart.Series)
{
    cs.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = true;
}

// Set the legend position to the top
chart.Legend.Position = LegendPositionType.Top;
```

---

# spire.xls csharp funnel chart
## create funnel chart using spire.xls
```csharp
//Add a new chart
var officeChart = sheet.Charts.Add();

//Set chart type as Funnel
officeChart.ChartType = ExcelChartType.Funnel;

//Set data range in the worksheet
officeChart.DataRange = sheet.Range["A1:B6"];

//Set the chart title
officeChart.ChartTitle = "Funnel";

//Formatting the legend and data label option
officeChart.HasLegend = false;
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
```

---

# spire.xls csharp histogram chart
## create histogram chart with spire.xls
```csharp
//Add a new chart
var officeChart = sheet.Charts.Add();

//Set chart type as histogram       
officeChart.ChartType = ExcelChartType.Histogram;

//Set data range in the worksheet   
officeChart.DataRange = sheet["A1:A15"];
officeChart.TopRow = 1;
officeChart.BottomRow = 19;
officeChart.LeftColumn = 4;
officeChart.RightColumn = 12;

//Category axis bin settings        
officeChart.PrimaryCategoryAxis.BinWidth = 8;

//Gap width settings
officeChart.Series[0].DataFormat.Options.GapWidth = 6;

//Set the chart title and axis title
officeChart.ChartTitle = "Height Data";
officeChart.PrimaryValueAxis.Title = "Number of students";
officeChart.PrimaryCategoryAxis.Title = "Height";

//Hiding the legend
officeChart.HasLegend = false;
```

---

# spire.xls csharp multi-level chart
## Create a multi-level category chart using Spire.XLS
```csharp
// Add a clustered bar chart to worksheet
Chart chart = sheet.Charts.Add(ExcelChartType.BarClustered);
chart.ChartTitle = "Value";
chart.PlotArea.Fill.FillType = ShapeFillType.NoFill;
chart.Legend.Delete();
chart.LeftColumn = 5;
chart.TopRow = 1;
chart.RightColumn = 14;

// Set the data source of series data
chart.DataRange = sheet.Range["C2:C9"];
chart.SeriesDataFromRange = false;
// Set the data source of category labels
ChartSerie serie = chart.Series[0];
serie.CategoryLabels = sheet.Range["A2:B9"];
// Show multi-level category labels
chart.PrimaryCategoryAxis.MultiLevelLable = true;
```

---

# spire.xls csharp pareto chart
## create a pareto chart with customization options
```csharp
// Add chart
Chart officeChart = sheet.Charts.Add();

// Set chart type as Pareto
officeChart.ChartType = ExcelChartType.Pareto;

// Set data range in the worksheet
officeChart.DataRange = sheet["A2:B8"];
officeChart.TopRow = 1;
officeChart.BottomRow = 19;
officeChart.LeftColumn = 4;
officeChart.RightColumn = 12;
officeChart.PrimaryCategoryAxis.IsBinningByCategory = true;

officeChart.PrimaryCategoryAxis.OverflowBinValue = 5;
officeChart.PrimaryCategoryAxis.UnderflowBinValue = 1;

// Formatting Pareto line
officeChart.Series[0].ParetoLineFormat.LineProperties.Color = Color.Blue;

// Gap width settings
officeChart.Series[0].DataFormat.Options.GapWidth = 6;

// Set the chart title
officeChart.ChartTitle = "Expenses";

// Hiding the legend
officeChart.HasLegend = false;
```

---

# spire.xls csharp pivot chart
## create pivot chart from pivot table
```csharp
//get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
//get the first pivot table in the worksheet
IPivotTable pivotTable = sheet.PivotTables[0];

//create a clustered column chart based on the pivot table
Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, pivotTable);

//set chart position
chart.TopRow = 12;
chart.LeftColumn = 1;
chart.RightColumn = 8;
chart.BottomRow = 30;
chart.ChartTitle = "Product";
chart.PrimaryCategoryAxis.MultiLevelLable = true;
```

---

# Spire.XLS C# Radar Chart Creation
## Code to create and configure a radar chart in Excel using Spire.XLS library
```csharp
//Add a new chart worksheet to workbook
Chart chart = sheet.Charts.Add();

//Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;

//Set region of chart data
chart.DataRange = sheet.Range["A1:C5"];
chart.SeriesDataFromRange = false;

// Set chart type
if (checkBox1.Checked)
{
    chart.ChartType = ExcelChartType.RadarFilled;
}
else
{
    chart.ChartType = ExcelChartType.Radar;
}

//Set chart title
chart.ChartTitle = "Sale market by region";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
// Set the visibility of plot area fill to false
chart.PlotArea.Fill.Visible = false;
// Set the position of the legend to corner
chart.Legend.Position = LegendPositionType.Corner;
```

---

# spire.xls csharp sunburst chart
## create sunburst chart in excel
```csharp
// Add chart
Chart officeChart = sheet.Charts.Add();

// Set chart type as Sunburst
officeChart.ChartType = ExcelChartType.SunBurst;

//Set data range in the worksheet
officeChart.DataRange = sheet["A1:D16"];
officeChart.TopRow = 1;
officeChart.BottomRow = 17;
officeChart.LeftColumn = 6;
officeChart.RightColumn = 14;

// Set the chart title
officeChart.ChartTitle = "Sales by quarter";

// Formatting data labels      
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;

// Hiding the legend
officeChart.HasLegend = false;
```

---

# spire.xls csharp treemap chart
## create TreeMap chart using Spire.XLS library
```csharp
// Add chart
Chart officeChart = sheet.Charts.Add();

// Set chart type as TreeMap
officeChart.ChartType = ExcelChartType.TreeMap;
 
// Set data range in the worksheet
officeChart.DataRange = sheet["A2:C11"];
officeChart.TopRow = 1;
officeChart.BottomRow = 19;
officeChart.LeftColumn = 4;
officeChart.RightColumn = 14;

// Set the chart title
officeChart.ChartTitle = "Area by countries";

// Set the Treemap label option
officeChart.Series[0].DataFormat.TreeMapLabelOption = ExcelTreeMapLabelOption.Banner;

// Formatting data labels      
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
```

---

# spire.xls csharp waterfall chart
## create and configure a waterfall chart in Excel
```csharp
// Add a new chart to the worksheet
var officeChart = sheet.Charts.Add();

// Set chart type as waterfall
officeChart.ChartType = ExcelChartType.WaterFall;

// Set data range for the chart from the worksheet
officeChart.DataRange = sheet["A2:B8"];

// Set chart position and size
officeChart.TopRow = 1;
officeChart.BottomRow = 19;
officeChart.LeftColumn = 4;
officeChart.RightColumn = 12;

// Set certain data points in the chart as totals
officeChart.Series[0].DataPoints[3].SetAsTotal = true;
officeChart.Series[0].DataPoints[6].SetAsTotal = true;

// Show connector lines between data points
officeChart.Series[0].Format.ShowConnectorLines = true;

// Set the chart title
officeChart.ChartTitle = "Waterfall Chart";

// Format data labels and legend options
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
officeChart.Legend.Position = LegendPositionType.Right;
```

---

# spire.xls csharp chart
## customize data markers in scatter chart
```csharp
//Create a Scatter-Markers chart based on the sample data
Chart chart = sheet.Charts.Add(ExcelChartType.ScatterMarkers);
chart.DataRange = sheet.Range["A1:B7"];
chart.PlotArea.Visible = false;
chart.SeriesDataFromRange = false;
chart.TopRow = 5;
chart.BottomRow = 22;
chart.LeftColumn = 4;
chart.RightColumn = 11;
chart.ChartTitle = "Chart with Markers";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 10;

//Format the markers in the chart by setting the background color, foreground color, type, size and transparency
Spire.Xls.Charts.ChartSerie cs1 = chart.Series[0];
cs1.DataFormat.MarkerBackgroundColor = Color.RoyalBlue;
cs1.DataFormat.MarkerForegroundColor = Color.WhiteSmoke;
cs1.DataFormat.MarkerSize = 7;
cs1.DataFormat.MarkerStyle = ChartMarkerType.PlusSign;
cs1.DataFormat.MarkerTransparencyValue = 0.8;

Spire.Xls.Charts.ChartSerie cs2 = chart.Series[1];
cs2.DataFormat.MarkerBackgroundColor = Color.Pink;
cs2.DataFormat.MarkerSize = 9;
cs2.DataFormat.MarkerStyle = ChartMarkerType.Triangle;
cs2.DataFormat.MarkerTransparencyValue = 0.9;
```

---

# spire.xls csharp chart data callout
## configure data callout settings for chart series
```csharp
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Get the first chart
Chart chart = sheet.Charts[0];

// Enable data labels and customize callout settings for each series in the chart
foreach (ChartSerie cs in chart.Series)
{
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasWedgeCallout = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = true;
}
```

---

# spire.xls csharp chart legend
## delete legend entries from excel chart
```csharp
// Get the chart
Chart chart = sheet.Charts[0];

//Delete the first and the second legend entries from the chart
chart.Legend.LegendEntries[0].Delete();
chart.Legend.LegendEntries[1].Delete();
```

---

# spire.xls csharp chart
## create chart with discontinuous data
```csharp
// Add a chart
Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered);
chart.SeriesDataFromRange = false;

// Set the position of chart
chart.LeftColumn = 1;
chart.TopRow = 10;
chart.RightColumn = 10;
chart.BottomRow = 24;

// Add a series
ChartSerie cs1 = (ChartSerie)chart.Series.Add();

// Set the name of the cs1
cs1.Name = sheet.Range["B1"].Value;

// Set discontinuous values for cs1
cs1.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"]);
cs1.Values = sheet.Range["B2:B3"].AddCombinedRange(sheet.Range["B5:B6"]).AddCombinedRange(sheet.Range["B8:B9"]);

//Set the chart type
cs1.SerieType = ExcelChartType.ColumnClustered;

// Add a series
ChartSerie cs2 = (ChartSerie)chart.Series.Add();
cs2.Name = sheet.Range["C1"].Value;
cs2.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"]);
cs2.Values = sheet.Range["C2:C3"].AddCombinedRange(sheet.Range["C5:C6"]).AddCombinedRange(sheet.Range["C8:C9"]);
cs2.SerieType = ExcelChartType.ColumnClustered;

// Set the chart title
chart.ChartTitle = "Chart";
chart.ChartTitleArea.Size = 20;
chart.ChartTitleArea.Color = Color.Black;

// Disable major grid lines on the primary value axis
chart.PrimaryValueAxis.HasMajorGridLines = false;
```

---

# spire.xls csharp chart
## edit line chart by adding new series
```csharp
// Get the line chart
Chart chart = sheet.Charts[0];

// Add a new series
ChartSerie cs = chart.Series.Add("Added");

// Set the values for the series
cs.Values = sheet.Range["I1:L1"];
```

---

# spire.xls csharp chart
## create exploded doughnut chart
```csharp
// Add a chart
Chart chart = sheet.Charts.Add();
chart.ChartType = ExcelChartType.DoughnutExploded;

// Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;

// Set region of chart data
chart.DataRange = sheet.Range["A1:B5"];
chart.SeriesDataFromRange = false;

// Chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

foreach (ChartSerie cs in chart.Series)
{
    // Enable varying colors for each data point
    cs.Format.Options.IsVaryColor = true;
    // Show data labels for data points
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
}

// Hide plot area fill
chart.PlotArea.Fill.Visible = false;

// Set legend position to the top
chart.Legend.Position = LegendPositionType.Top;
```

---

# Spire.XLS C# Trendline Extraction
## Extract trendline equation from an Excel chart
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load an Excel file
workbook.LoadFromFile("filePath");

// Get the chart from the first worksheet
Chart chart = workbook.Worksheets[0].Charts[0];

// Get the trendline of the chart and then extract the equation of the trendline
IChartTrendLine trendLine = chart.Series[1].TrendLines[0];
string formula = trendLine.Formula;
StringBuilder sb = new StringBuilder();
sb.AppendLine("The equation is: " + formula);
```

---

# Spire.XLS C# Chart Element Picture Fill
## Fill chart elements with custom pictures
```csharp
//Get the first chart
Chart chart = ws.Charts[0];

// A. Fill chart area with image
chart.ChartArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\background.png"), "None");
chart.PlotArea.Fill.Transparency = 0.9;

//// B.Fill plot area with image
//chart.PlotArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\background.png"), "None");
```

---

# spire.xls csharp fill picture for chart marker
## fill picture for chart marker in excel
```csharp
// Get the first chart from the worksheet.
Chart chart = worksheet.Charts[0];

// Set the line color of series 1 to yellow.
chart.Series[0].Format.LineProperties.Color = Color.Yellow;

// Set the marker style of series 1 to picture.
chart.Series[0].Format.MarkerStyle = ChartMarkerType.Picture;

// Get the marker fill for series 1.
IShapeFill markerFill1 = chart.Series[0].DataFormat.MarkerFill;

// Set the custom picture for the marker fill of series 1.
markerFill1.CustomPicture(imageFile);

// Get the marker fill for series 2.
IShapeFill markerFill2 = chart.Series[1].DataFormat.MarkerFill;

// Set the line color of series 2 to red.
chart.Series[1].Format.LineProperties.Color = Color.Red;

// Set the texture of the marker fill for series 2 to granite.
markerFill2.Texture = GradientTextureType.Granite;

// Set the line color of series 1 to blue.
chart.Series[0].Format.LineProperties.Color = Color.Blue;

// Get the marker fill for series 3.
IShapeFill markerFill3 = chart.Series[2].DataFormat.MarkerFill;

// Set the pattern of the marker fill for series 3 to 10% gradient
markerFill3.Pattern = GradientPatternType.Pat10Percent;

// Set the foreground color of the marker fill for series 3 to light gray.
markerFill3.ForeColor = Color.LightGray;

// Set the background color of the marker fill for series 3 to orange.
markerFill3.BackColor = Color.Orange;
```

---

# spire.xls csharp chart axis formatting
## format chart axis in excel
```csharp
//Add a chart
Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered);
chart.DataRange = sheet.Range["B1:B9"];
chart.SeriesDataFromRange = false;
chart.PlotArea.Visible = false;
chart.TopRow = 10;
chart.BottomRow = 28;
chart.LeftColumn = 2;
chart.RightColumn = 10;
chart.ChartTitle = "Chart with Customized Axis";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
Spire.Xls.Charts.ChartSerie cs1 = chart.Series[0];
cs1.CategoryLabels = sheet.Range["A2:A9"];

//Format axis
chart.PrimaryValueAxis.MajorUnit = 8;
chart.PrimaryValueAxis.MinorUnit = 2;
chart.PrimaryValueAxis.MaxValue = 50;
chart.PrimaryValueAxis.MinValue = 0;
chart.PrimaryValueAxis.IsReverseOrder = false;
chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkOutside;
chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkInside;
chart.PrimaryValueAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionNextToAxis;
chart.PrimaryValueAxis.CrossesAt = 0;

//Set NumberFormat
chart.PrimaryValueAxis.NumberFormat = "$#,##0";
chart.PrimaryValueAxis.IsSourceLinked = false;

ChartSerie serie = chart.Series[0];

foreach (ChartDataPoint dataPoint in serie.DataPoints)
{
    //Format Series
    dataPoint.DataFormat.Fill.FillType = ShapeFillType.SolidColor;
    dataPoint.DataFormat.Fill.ForeColor = Color.LightGreen;

    //Set transparency
    dataPoint.DataFormat.Fill.Transparency = 0.3;           
}
```

---

# spire.xls csharp gauge chart
## create a gauge chart using doughnut and pie charts
```csharp
// Create a Workbook
Workbook workbook = new Workbook();

// Get the first sheet and set its name
Worksheet sheet = workbook.Worksheets[0];
sheet.Name = "Gauge Chart";

// Add a Doughnut chart
Chart chart = sheet.Charts.Add(ExcelChartType.Doughnut);
chart.DataRange = sheet.Range["A1:A5"];
chart.SeriesDataFromRange = false;
chart.HasLegend = true;

// Set the position of chart
chart.LeftColumn = 2;
chart.TopRow = 7;
chart.RightColumn = 9;
chart.BottomRow = 25;

// Get the series 1
ChartSerie cs1 = (ChartSerie)chart.Series["Value"];
cs1.Format.Options.DoughnutHoleSize = 60;
cs1.DataFormat.Options.FirstSliceAngle = 270;

// Set the fill color
cs1.DataPoints[0].DataFormat.Fill.ForeColor = Color.Yellow;
cs1.DataPoints[1].DataFormat.Fill.ForeColor = Color.PaleVioletRed;
cs1.DataPoints[2].DataFormat.Fill.ForeColor = Color.DarkViolet;
cs1.DataPoints[3].DataFormat.Fill.Visible = false;

// Add a series with pie chart
ChartSerie cs2 = (ChartSerie)chart.Series.Add("Pointer", ExcelChartType.Pie);

// Set the value
cs2.Values = sheet.Range["D2:D4"];
cs2.UsePrimaryAxis = false;
cs2.DataPoints[0].DataLabels.HasValue = true;
cs2.DataFormat.Options.FirstSliceAngle = 270;
cs2.DataPoints[0].DataFormat.Fill.Visible = false;
cs2.DataPoints[1].DataFormat.Fill.FillType = ShapeFillType.SolidColor;
cs2.DataPoints[1].DataFormat.Fill.ForeColor = Color.Black;
cs2.DataPoints[2].DataFormat.Fill.Visible = false;
```

---

# Spire.XLS C# Chart Category Labels
## Extract category labels from a chart in Excel
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Get the first chart
Chart chart = sheet.Charts[0];

// Get the cell range of the category labels
CellRange cr = chart.PrimaryCategoryAxis.CategoryLabels;
foreach (var cell in cr)
{
    // Process each category label
    string categoryLabel = cell.Value;
}
```

---

# spire.xls csharp get chart data
## retrieve data point values from an Excel chart
```csharp
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Get the chart
Chart chart = sheet.Charts[0];

// Get the first series of the chart
ChartSerie cs = chart.Series[0];

foreach (CellRange cr in cs.Values)
{
    // Get the data point value
    string value = cr.Value;
}
```

---

# spire.xls csharp chart worksheet
## get worksheet containing chart
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Load the Excel document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartToImage.xlsx");

//Access first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];

//Access the first chart inside this worksheet
Chart chart = worksheet.Charts[0];

//Get its worksheet
Worksheet wSheet = chart.Worksheet as Worksheet;
```

---

# spire.xls csharp chart
## hide major gridlines in excel chart
```csharp
//Get the chart
Chart chart = sheet.Charts[0];

//Hide major gridlines
chart.PrimaryValueAxis.HasMajorGridLines = false;
```

---

# spire.xls csharp line chart
## create and customize line chart in excel
```csharp
// Add a chart
Chart chart = sheet.Charts.Add();

// Set chart type based on selection
if (isLine3D)
{
    chart.ChartType = ExcelChartType.Line3D;
}
else
{
    chart.ChartType = ExcelChartType.Line;
}

// Set region of chart data
chart.DataRange = sheet.Range["A1:E5"];

// Set position of the chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;

// Set chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

// Customize primary category axis (x-axis)
chart.PrimaryCategoryAxis.Title = "Month";
chart.PrimaryCategoryAxis.Font.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

// Customize primary value axis (y-axis)
chart.PrimaryValueAxis.Title = "Sales (in Dollars)";
chart.PrimaryValueAxis.HasMajorGridLines = false;
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;
chart.PrimaryValueAxis.MinValue = 1000;
chart.PrimaryValueAxis.TitleArea.IsBold = true;

// Customize series settings
foreach (ChartSerie cs in chart.Series)
{
    // Enable varying colors for each data point
    cs.Format.Options.IsVaryColor = true; 
    // Show data labels for data points
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true; 
    // Set marker style for data points
    if (!isLine3D)
        cs.DataFormat.MarkerStyle = ChartMarkerType.Circle; 
}

// Hide plot area fill
chart.PlotArea.Fill.Visible = false;

// Set legend position to the top
chart.Legend.Position = LegendPositionType.Top; 
```

---

# spire.xls csharp line chart droplines
## add drop lines to line chart
```csharp
// Get the first chart
Chart chart = worksheet.Charts[0];

// Add a drop lines to the first series
chart.Series[0].HasDroplines = true;
```

---

# spire.xls csharp pie chart
## create pie chart with Spire.XLS library
```csharp
//Add a chart
Chart chart = null;
if (is3DPie)
{
    chart = sheet.Charts.Add(ExcelChartType.Pie3D);
}
else
{
    chart = sheet.Charts.Add(ExcelChartType.Pie);
}

//Set region of chart data
chart.DataRange = sheet.Range["B2:B5"];
chart.SeriesDataFromRange = false;

//Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 9;
chart.BottomRow = 25;

//Chart title
chart.ChartTitle = "Sales by year";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

ChartSerie cs = chart.Series[0];
cs.CategoryLabels = sheet.Range["A2:A5"];
cs.Values = sheet.Range["B2:B5"];
cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;

// Hide plot area fill
chart.PlotArea.Fill.Visible = false;
```

---

# spire.xls csharp pyramid chart
## create pyramid column chart with customization options
```csharp
// Add a chart
Chart chart = sheet.Charts.Add();

// Set region of chart data
chart.DataRange = sheet.Range["B2:B5"];
chart.SeriesDataFromRange = false;

// Set position of the chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;

// Set chart type based on checkbox selection
if (is3DChecked)
{
    chart.ChartType = ExcelChartType.Pyramid3DClustered;
}
else
{
    chart.ChartType = ExcelChartType.PyramidClustered;
}

// Set chart title
chart.ChartTitle = "Sales by year";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

// Customize primary category axis (x-axis)
chart.PrimaryCategoryAxis.Title = "Year";
chart.PrimaryCategoryAxis.Font.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

// Customize primary value axis (y-axis)
chart.PrimaryValueAxis.Title = "Sales (in Dollars)";
chart.PrimaryValueAxis.HasMajorGridLines = false;
chart.PrimaryValueAxis.MinValue = 1000;
chart.PrimaryValueAxis.TitleArea.IsBold = true;
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;

// Customize series settings
ChartSerie cs = chart.Series[0];
// Set category labels for the series
cs.CategoryLabels = sheet.Range["A2:A5"];
// Enable varying colors for each data point
cs.Format.Options.IsVaryColor = true; 
// Set legend position to the top
chart.Legend.Position = LegendPositionType.Top; 
```

---

# Spire.XLS C# Chart Removal
## Remove a chart from an Excel worksheet
```csharp
//Get the first chart from the worksheet
IChartShape chart = worksheet.Charts[0];

//Remove the chart
chart.Remove();
```

---

# spire.xls csharp chart
## resize and move chart in excel
```csharp
//Get the chart from the first worksheet
Worksheet sheet = workbook.Worksheets[0];
Chart chart = sheet.Charts[0];

//Set position of the chart
chart.LeftColumn = 5;
chart.TopRow = 1;

//Resize the chart
chart.Width = 500;
chart.Height = 350;
```

---

# spire.xls csharp chart rich text
## Set rich text for data labels in Excel chart
```csharp
//Get first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];

//Get the first chart inside this worksheet
Chart chart = worksheet.Charts[0];

//Get the first datalabel of the first series 
ChartDataLabels datalabel = chart.Series[0].DataPoints[0].DataLabels;

//Set the text
datalabel.Text = "Rich Text Label";

//Show the value
chart.Series[0].DataPoints[0].DataLabels.HasValue = true;

//Set styles for the text
chart.Series[0].DataPoints[0].DataLabels.Color = Color.Red;
chart.Series[0].DataPoints[0].DataLabels.IsBold = true;
```

---

# spire.xls csharp 3d chart rotation
## rotate 3d chart in excel
```csharp
//Get the chart from the first worksheet
Worksheet sheet = workbook.Worksheets[0];
Chart chart = sheet.Charts[0];

//X rotation:
chart.Rotation = 30;
//Y rotation:
chart.Elevation = 20;
```

---

# spire.xls csharp chart data labels
## Set and format data labels for a chart in Excel using Spire.XLS
```csharp
// Add a line chart with markers
Chart chart = sheet.Charts.Add(ExcelChartType.LineMarkers);

// Set chart data range and position
chart.DataRange = sheet.Range["B1:B7"];
chart.PlotArea.Visible = false;
chart.SeriesDataFromRange = false;
chart.TopRow = 5;
chart.BottomRow = 26;
chart.LeftColumn = 2;
chart.RightColumn = 11;

// Set chart title
chart.ChartTitle = "Data Labels Demo";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

// Customize series settings
Spire.Xls.Charts.ChartSerie cs1 = chart.Series[0];
// Set category labels for the series
cs1.CategoryLabels = sheet.Range["A2:A7"]; 

// Customize data label settings for default data point
cs1.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
cs1.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = false;
cs1.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = false;
cs1.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = true;
cs1.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = true;
cs1.DataPoints.DefaultDataPoint.DataLabels.Delimiter = ". ";
cs1.DataPoints.DefaultDataPoint.DataLabels.Size = 9;
cs1.DataPoints.DefaultDataPoint.DataLabels.Color = Color.Red;
cs1.DataPoints.DefaultDataPoint.DataLabels.FontName = "Calibri";
cs1.DataPoints.DefaultDataPoint.DataLabels.Position = DataLabelPositionType.Center;
```

---

# spire.xls csharp chart border
## Set border color and style for Excel chart series
```csharp
//Get the first worksheet from workbook and then get the first chart from the worksheet
Worksheet ws = workbook.Worksheets[0];
Chart chart = ws.Charts[0];

//Set CustomLineWeight property for Series line
(chart.Series[0].DataPoints[0].DataFormat.LineProperties as XlsChartBorder).CustomLineWeight = 2.5f;
//Set color property for Series line
(chart.Series[0].DataPoints[0].DataFormat.LineProperties as XlsChartBorder).Color = Color.Red;
```

---

# spire.xls csharp marker border
## set border width of chart markers
```csharp
// Get the chart from the first worksheet
Chart chart = workbook.Worksheets[0].Charts[0];

// Set marker border width for series 1
chart.Series[0].DataFormat.MarkerBorderWidth = 1.5; 

// Set marker border width for series 2
chart.Series[1].DataFormat.MarkerBorderWidth = 2.5; 
```

---

# spire.xls csharp chart styling
## Set chart background color
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Get the first worksheet from workbook and then get the first chart from the worksheet
Worksheet ws = workbook.Worksheets[0];
Chart chart = ws.Charts[0];

//Set background color
chart.ChartArea.ForeGroundColor = System.Drawing.Color.LightYellow;
```

---

# spire.xls csharp chart font
## set font for chart legend and data labels
```csharp
//Get the first worksheet from workbook
Worksheet ws = workbook.Worksheets[0];
Chart chart = ws.Charts[0];

//Create a font with specified size and color
ExcelFont font = workbook.CreateFont();
font.Size = 14.0;
font.Color = Color.Red;

//Apply the font to chart Legend
chart.Legend.TextArea.SetFont(font);

//Apply the font to chart DataLabel
foreach (ChartSerie cs in chart.Series)
{
    cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font);
}
```

---

# spire.xls csharp chart font formatting
## set font for chart title and axis
```csharp
//Set font for chart title and chart axis
Worksheet worksheet = workbook.Worksheets[0];
Chart chart = worksheet.Charts[0];

//Format the font for the chart title
chart.ChartTitleArea.Color = Color.Blue;
chart.ChartTitleArea.Size = 20.0;
chart.ChartTitleArea.FontName = "Arial";

//Format the font for the chart Axis
chart.PrimaryValueAxis.Font.Color = Color.Gold;
chart.PrimaryValueAxis.Font.Size = 10.0;
chart.PrimaryCategoryAxis.Font.FontName = "Arial";
chart.PrimaryCategoryAxis.Font.Color = Color.Red;
chart.PrimaryCategoryAxis.Font.Size = 20.0;
```

---

# spire.xls chart legend background
## set legend background color in excel chart
```csharp
// Get the chart from the worksheet
Chart chart = ws.Charts[0];

// Access the legend frame format and set the background color
XlsChartFrameFormat x = chart.Legend.FrameFormat as XlsChartFrameFormat;
x.Fill.FillType = ShapeFillType.SolidColor;
x.ForeGroundColor = Color.SkyBlue;
```

---

# spire.xls csharp trendline
## set number format of trendline
```csharp
//Get the chart from the first worksheet
Chart chart = workbook.Worksheets[0].Charts[0];

//Get the trendline of the chart and then extract the equation of the trendline
IChartTrendLine trendLine = chart.Series[1].TrendLines[0];

//Set the number format of trendLine to "#,##0.00"
trendLine.DataLabel.NumberFormat = "#,##0.00";
```

---

# spire.xls csharp chart leader lines
## enable leader lines for chart data labels
```csharp
// Add a stacked bar chart
Chart chart = sheet.Charts.Add(ExcelChartType.BarStacked);
chart.DataRange = sheet.Range["A1:C3"];
chart.TopRow = 4;
chart.LeftColumn = 2;
chart.Width = 450;
chart.Height = 300;

// Enable data labels with leader lines for each series
foreach (ChartSerie cs in chart.Series)
{
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = true;
}
```

---

# spire.xls csharp sparkline
## create sparkline in excel worksheet
```csharp
//Add sparkline
SparklineGroup sparklineGroup = sheet.SparklineGroups.AddGroup(SparklineType.Line);
SparklineCollection sparklines = sparklineGroup.Add();
sparklines.Add(sheet["A2:D2"], sheet["E2"]);
sparklines.Add(sheet["A3:D3"], sheet["E3"]);
sparklines.Add(sheet["A4:D4"], sheet["E4"]);
sparklines.Add(sheet["A5:D5"], sheet["E5"]);
sparklines.Add(sheet["A6:D6"], sheet["E6"]);
sparklines.Add(sheet["A7:D7"], sheet["E7"]);
sparklines.Add(sheet["A8:D8"], sheet["E8"]);
sparklines.Add(sheet["A9:D9"], sheet["E9"]);
sparklines.Add(sheet["A10:D10"], sheet["E10"]);
sparklines.Add(sheet["A11:D11"], sheet["E11"]);
```

---

# spire.xls csharp shapes
## add arrow lines to excel worksheet
```csharp
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Add a Double Arrow and fill the line with solid color.
var line = sheet.TypedLines.AddLine();
line.Top = 10;
line.Left = 20;
line.Width = 100;
line.Height = 0;
line.Color = Color.Blue;
line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow;
line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;

//Add an Arrow and fill the line with solid color.
var line_1 = sheet.TypedLines.AddLine();
line_1.Top = 50;
line_1.Left = 30;
line_1.Width = 100;
line_1.Height = 100;
line_1.Color = Color.Red;
line_1.BeginArrowHeadStyle = ShapeArrowStyleType.LineNoArrow;
line_1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;

//Add an Elbow Arrow Connector.
Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape line3 = sheet.TypedLines.AddLine() as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
line3.LineShapeType = LineShapeType.ElbowLine;
line3.Width = 30;
line3.Height = 50;
line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;
line3.Top = 100;
line3.Left = 50;

//Add an Elbow Double-Arrow Connector.
Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape line2 = sheet.TypedLines.AddLine() as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
line2.LineShapeType = LineShapeType.ElbowLine;
line2.Width = 50;
line2.Height = 50;
line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;
line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow;
line2.Left = 120;
line2.Top = 100;

//Add a Curved Arrow Connector.
line3 = sheet.TypedLines.AddLine() as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
line3.LineShapeType = LineShapeType.CurveLine;
line3.Width = 30;
line3.Height = 50;
line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen;
line3.Top = 100;
line3.Left = 200;

//Add a Curved Double-Arrow Connector.
line2 = sheet.TypedLines.AddLine() as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
line2.LineShapeType = LineShapeType.CurveLine;
line2.Width = 30;
line2.Height = 50;
line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen;
line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen;
line2.Left = 250;
line2.Top = 100;
```

---

# spire.xls csharp line shapes
## add different line shapes to excel worksheet
```csharp
//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Add shape line1
ILineShape line1 = sheet.Lines.AddLine(10, 2, 200, 1, LineShapeType.Line);
//Set dash style type
line1.DashStyle = ShapeDashLineStyleType.Solid;
//Set color
line1.Color = Color.CadetBlue;
//Set weight
line1.Weight = 2f;
//Set end arrow style type
line1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;

//Add shape line2
ILineShape line2 = sheet.Lines.AddLine(12, 2, 200, 1, LineShapeType.CurveLine);
line2.DashStyle = ShapeDashLineStyleType.Dotted;
line2.Color = Color.OrangeRed;
line2.Weight = 2f;

//Add shape line3
ILineShape line3 = sheet.Lines.AddLine(14, 2, 200, 1, LineShapeType.ElbowLine);
line3.DashStyle = ShapeDashLineStyleType.DashDotDot;
line3.Color = Color.Purple;
line3.Weight = 2f;

//Add shape line4
ILineShape line4 = sheet.Lines.AddLine(16, 2, 200, 1, LineShapeType.LineInv);
line4.DashStyle = ShapeDashLineStyleType.Dashed;
line4.Color = Color.Green;
line4.Weight = 2f;
```

---

# spire.xls csharp add oval shape
## Add oval shapes to Excel worksheet with different fill styles
```csharp
//Add oval shape1
IOvalShape ovalShape1 = sheet.OvalShapes.AddOval(11, 2, 100, 100);
ovalShape1.Line.Weight = 0;
//Fill shape with solid color
ovalShape1.Fill.FillType = ShapeFillType.SolidColor;
ovalShape1.Fill.ForeColor = Color.DarkCyan;

//Add oval shape2
IOvalShape ovalShape2 = sheet.OvalShapes.AddOval(11, 5, 100, 100);
ovalShape2.Line.Weight = 1;
//Fill shape with picture
ovalShape2.Line.DashStyle = ShapeDashLineStyleType.Solid;
ovalShape2.Fill.CustomPicture(@"..\..\..\..\..\..\Data\logo.png");
```

---

# spire.xls csharp shapes
## add rectangle shapes to excel worksheet
```csharp
//Add rectangle shape 1------Rect
IRectangleShape rect1=sheet.RectangleShapes.AddRectangle(11, 2, 60, 100, RectangleShapeType.Rect);
rect1.Line.Weight = 1;
//Fill shape with solid color
rect1.Fill.FillType = ShapeFillType.SolidColor;
rect1.Fill.ForeColor = Color.DarkGreen;

//Add rectangle shape 2------RoundRect
IRectangleShape rect2 = sheet.RectangleShapes.AddRectangle(11, 5, 60, 100, RectangleShapeType.RoundRect);
rect2.Line.Weight = 1;
rect2.Fill.FillType = ShapeFillType.SolidColor;
rect2.Fill.ForeColor = Color.DarkCyan;
```

---

# spire.xls csharp spinner control
## add spinner control to excel worksheet
```csharp
//Set text for range C11
sheet.Range["C11"].Text = "Value:";
sheet.Range["C11"].Style.Font.IsBold = true;

//Set value for range B10
sheet.Range["C12"].Value2 = 0;

//Add spinner control
ISpinnerShape spinner = sheet.SpinnerShapes.AddSpinner(12, 4, 20, 20);
spinner.LinkedCell = sheet.Range["C12"];
spinner.Min = 0;
spinner.Max = 100;
spinner.IncrementalChange = 5;
spinner.Display3DShading = true;
```

---

# spire.xls csharp arrow polyline
## adjust arrow polyline position in excel
```csharp
// Draw an elbow arrow
XlsLineShape line = worksheet.TypedLines.AddLine(5, 5, 100, 100, LineShapeType.ElbowLine) as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
line.EndArrowHeadStyle = ShapeArrowStyleType.LineNoArrow;
line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow;
GeomertyAdjustValue ad = line.ShapeAdjustValues.AddAdjustValue(GeomertyAdjustValueFormulaType.LiteralValue);

// When the parameter value is less than 0, the focus of the line is on the left side of the left point, when it is equal to 0, the position is the same as the left point, it is equal to 50 in the middle of the graph, and when it is equal to 100, it is the same as the right point.
ad.SetFormulaParameter(new double[] {-50});
```

---

# Spire.XLS C# Shapes to Images
## Convert Excel shapes to images
```csharp
//Get the first worksheet
Worksheet worksheet = workbook.Worksheets[0];

// Save all shape to images
SaveShapeTypeOption shapelist = new SaveShapeTypeOption();
shapelist.SaveAll = true;
List<Bitmap> images = worksheet.SaveShapesToImage(shapelist);
```

---

# Spire.XLS C# Shape Copying
## Copy shapes between worksheets in Excel
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Create line shape
var line = sheet.TypedLines.AddLine();
line.Top = 50;
line.Left = 30;
line.Width = 30;
line.Height = 50;
line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowDiamond;
line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;

// Get the second worksheet
Worksheet CopyShapes = workbook.Worksheets[1];

// Copy the line into other sheet
CopyShapes.TypedLines.AddCopy(line);

// Create a button and then copy into other sheet
var button = sheet.TypedRadioButtons.Add(5, 5, 20, 20);
CopyShapes.TypedRadioButtons.AddCopy(button);

// Create a textbox and then copy into other sheet
var textbox = sheet.TypedTextBoxes.AddTextBox(5, 7, 50, 100);
CopyShapes.TypedTextBoxes.AddCopy(textbox);

// Create a checkbox and then copy into other sheet
var checkbox = sheet.TypedCheckBoxes.AddCheckBox(10, 1, 20, 20);
CopyShapes.TypedCheckBoxes.AddCopy(checkbox);

// Create a comboboxes and then copy into other sheet
var ComboBoxes = sheet.TypedComboBoxes.AddComboBox(10, 5, 30, 30);
CopyShapes.TypedComboBoxes.AddCopy(ComboBoxes);
```

---

# spire.xls csharp delete shapes
## delete all shapes from Excel worksheet
```csharp
//Delete all shapes in the worksheet
for (int i = sheet.PrstGeomShapes.Count-1; i >= 0; i--)
{
    sheet.PrstGeomShapes[i].Remove();
}
```

---

# spire.xls csharp delete shape
## delete a particular shape from Excel worksheet
```csharp
//Delete the first shape in the worksheet
sheet.PrstGeomShapes[0].Remove();
```

---

# spire.xls csharp drawing lines
## draw lines through two points in excel using relative and absolute positions
```csharp
// Draw a line according to relative position
XlsLineShape line1 = worksheet.TypedLines.AddLine() as XlsLineShape;
line1.LeftColumn = 3;
line1.TopRow = 3;
line1.LeftColumnOffset = 0;
line1.TopRowOffset = 0; 

line1.RightColumn = 4; 
line1.BottomRow = 5; 
line1.RightColumnOffset = 0;
line1.BottomRowOffset = 0; 

// Draw a line according to absolute position(pixels)
XlsLineShape line2 = worksheet.TypedLines.AddLine() as XlsLineShape;
line2.StartPoint = new Point(30, 50);
line2.EndPoint = new Point(20, 80);
```

---

# Extract Text and Image from Excel Shapes
## This code demonstrates how to extract text and image from shapes in an Excel worksheet using Spire.XLS library.
```csharp
//Extract text from the first shape.
IPrstGeomShape shape1 = sheet.PrstGeomShapes[2];
String s = shape1.Text;
StringBuilder sb = new StringBuilder();
sb.AppendLine("The text in the third shape is: " + s);

//Extract image from the second shape.
IPrstGeomShape shape2 = sheet.PrstGeomShapes[1];
Image image = shape2.Fill.Picture;
```

---

# Spire.XLS C# Shape Texture Fill
## Fill a shape with a picture as texture in Excel
```csharp
// Get the first shape
IPrstGeomShape shape = sheet.PrstGeomShapes[0];

// Fill shape with texture
shape.Fill.FillType = ShapeFillType.Texture;

// Custom texture with picture
shape.Fill.CustomTexture(@"..\..\..\..\..\..\Data\logo.png");

// Tile picture as texture 
shape.Fill.Tile = true;
```

---

# spire.xls csharp get shape linked cell range
## retrieve cell ranges linked to shapes in excel worksheet
```csharp
// Get the first worksheet from the workbook.
Worksheet sheet = workbook.Worksheets[0];

// Get the collection of preset geometric shapes in the sheet.
PrstGeomShapeCollection prstGeomShapeCollection = sheet.PrstGeomShapes;

// Get a specific shape by its name.
IPrstGeomShape shape = prstGeomShapeCollection["Yesterday"];

// Get the range address of the cell linked to the shape.
string cellAddress = shape.LinkedCell.RangeAddress;

// Get another shape by its name.
shape = prstGeomShapeCollection["NewShapes"];

// Get the range address of the cell linked to the shape.
cellAddress = shape.LinkedCell.RangeAddress;
```

---

# Spire.XLS C# Shape Grouping
## Group multiple shapes in an Excel worksheet
```csharp
// Add shapes to the worksheet
IPrstGeomShape shape1 = worksheet.PrstGeomShapes.AddPrstGeomShape(1, 3, 50, 50, PrstGeomShapeType.RoundRect);
IPrstGeomShape shape2 = worksheet.PrstGeomShapes.AddPrstGeomShape(5, 3, 50, 50, PrstGeomShapeType.Triangle);

// Group the shapes
GroupShapeCollection groupShapeCollection = worksheet.GroupShapeCollection;
groupShapeCollection.Group(new Spire.Xls.Core.IShape[] { shape1, shape2 });
```

---

# spire.xls csharp group shapes to image
## convert grouped shapes in excel worksheet to images
```csharp
// Save to image
SaveShapeTypeOption saveShapeTypeOption = new SaveShapeTypeOption();
saveShapeTypeOption.SaveGroupShape = true;
List<Bitmap> images = worksheet.SaveShapesToImage(saveShapeTypeOption);
```

---

# Spire.XLS C# Shape Visibility
## Hide or unhide shapes in Excel worksheet
```csharp
//Hide the second shape in the worksheet
sheet.PrstGeomShapes[1].Visible = false;

//Show the second shape in the worksheet
//sheet.PrstGeomShapes[1].Visible = true;
```

---

# spire.xls csharp shapes
## insert different shapes with various fill types into excel sheet
```csharp
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Add a triangle shape.
IPrstGeomShape triangle = sheet.PrstGeomShapes.AddPrstGeomShape(2, 2, 100, 100, PrstGeomShapeType.Triangle);
//Fill the triangle with solid color.
triangle.Fill.ForeColor = Color.Yellow;
triangle.Fill.FillType = ShapeFillType.SolidColor;

//Add a heart shape.
IPrstGeomShape heart = sheet.PrstGeomShapes.AddPrstGeomShape(2, 5, 100, 100, PrstGeomShapeType.Heart);          
//Fill the heart with gradient color.
heart.Fill.ForeColor = Color.Red;
heart.Fill.FillType = ShapeFillType.Gradient;

//Add an arrow shape with default color.
IPrstGeomShape arrow = sheet.PrstGeomShapes.AddPrstGeomShape(10, 2, 100, 100, PrstGeomShapeType.CurvedRightArrow);

//Add a cloud shape.
IPrstGeomShape cloud = sheet.PrstGeomShapes.AddPrstGeomShape(10, 5, 100, 100, PrstGeomShapeType.Cloud);
//Fill the cloud with picture
cloud.Fill.FillType = ShapeFillType.Picture;
```

---

# Spire.XLS C# Shape Text Alignment
## Set middle-centered text in an Excel shape
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Add a rectangle shape to the worksheet
IPrstGeomShape rect = sheet.PrstGeomShapes.AddPrstGeomShape(8, 2, 300, 300, PrstGeomShapeType.Rect);

// Set the fill color of the rectangle to white (solid color)
rect.Fill.ForeColor = Color.White;
rect.Fill.FillType = ShapeFillType.SolidColor;

// Set the text content of the rectangle
rect.Text = "E-iceblue";

// Set the vertical alignment of the text to middle-centered
rect.TextVerticalAlignment = ExcelVerticalAlignment.MiddleCentered;
```

---

# spire.xls csharp shape shadow
## modify shadow style for shape
```csharp
//Get the third shape from the worksheet.
IPrstGeomShape shape = sheet.PrstGeomShapes[2];

//Set the shadow style for the shape.
shape.Shadow.Angle = 90;
shape.Shadow.Transparency = 30;
shape.Shadow.Distance = 10;
shape.Shadow.Size = 130;
shape.Shadow.Color = Color.Yellow;
shape.Shadow.Blur = 30;
shape.Shadow.HasCustomStyle = true;
```

---

# spire.xls csharp shape shadow
## set shadow style for shape
```csharp
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Add an ellipse shape.
IPrstGeomShape ellipse = sheet.PrstGeomShapes.AddPrstGeomShape(5, 5, 150, 100, PrstGeomShapeType.Ellipse);

//Set the shadow style for the ellipse.
ellipse.Shadow.Angle = 90;
ellipse.Shadow.Distance = 10;
ellipse.Shadow.Size = 150;
ellipse.Shadow.Color = Color.Gray;
ellipse.Shadow.Blur = 30;
ellipse.Shadow.Transparency = 1;
ellipse.Shadow.HasCustomStyle = true;
```

---

# spire.xls csharp shape order
## change the layer order of shapes in excel worksheets
```csharp
//Bring the picture forward one level
workbook.Worksheets[0].Pictures[0].ChangeLayer(ShapeLayerChangeType.BringForward);

//Bring the image in front of all other objects
workbook.Worksheets[1].Pictures[0].ChangeLayer(ShapeLayerChangeType.BringToFront);

//Send the shape back one level
XlsShape shape = workbook.Worksheets[2].PrstGeomShapes[1] as XlsShape;
shape.ChangeLayer(ShapeLayerChangeType.SendBackward);

//Send the shape behind all other objects
shape = workbook.Worksheets[3].PrstGeomShapes[1] as XlsShape;
shape.ChangeLayer(ShapeLayerChangeType.SendToBack);
```

---

# spire.xls csharp shape to image
## convert excel shape to image
```csharp
//Get the first shape from the worksheet
XlsShape shape = sheet1.PrstGeomShapes[0] as XlsShape;

//Save the shape to a image
Image img = shape.SaveToImage();
```

---

# spire.xls csharp shape conversion
## convert Excel shapes to images with options
```csharp
// Convert shapes to images
SaveShapeTypeOption shapelist = new SaveShapeTypeOption();

// Set the option to save all shapes in the worksheet to images
shapelist.SaveAll = true;

// Save the shapes in the worksheet as images and store them in a dictionary
Dictionary<IShape, Bitmap> images = sheet.SaveAndGetShapesToImage(shapelist);

// Iterate over each shape-image pair in the dictionary
foreach (KeyValuePair<IShape, Bitmap> pair in images)
{
    // Get the shape and image from the pair
    IShape shape = pair.Key;
    Bitmap bitmap = pair.Value;

    // Generate a unique image file name based on shape properties
    string imageFileName = shape.Name + "_" + shape.Height + "_" + shape.Width + "_" + shape.ShapeType + ".png";

    // Save the bitmap as an image file with the generated name
    bitmap.Save(imageFileName);
}
```

---

# spire.xls csharp apply built-in styles
## apply built-in title style to excel cells
```csharp
//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

//Apply title style
sheet.Range["A1:J1"].BuiltInStyle = BuiltInStyles.Title;
```

---

# spire.xls csharp conditional formatting
## apply color scales to data range in Excel
```csharp
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Insert data to cell range from A1 to C4.
sheet.Range["A1"].NumberValue = 582;
sheet.Range["A2"].NumberValue = 234;
sheet.Range["A3"].NumberValue = 314;
sheet.Range["A4"].NumberValue = 50;
sheet.Range["B1"].NumberValue = 150;
sheet.Range["B2"].NumberValue = 894;
sheet.Range["B3"].NumberValue = 560;
sheet.Range["B4"].NumberValue = 900;
sheet.Range["C1"].NumberValue = 134;
sheet.Range["C2"].NumberValue = 700;
sheet.Range["C3"].NumberValue = 920;
sheet.Range["C4"].NumberValue = 450;
sheet.AllocatedRange.RowHeight = 15;
sheet.AllocatedRange.ColumnWidth = 17;

//Add color scales.
XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.AllocatedRange);
IConditionalFormat format = xcfs.AddCondition();
format.FormatType = ConditionalFormatType.ColorScale;
```

---

# spire.xls csharp conditional formatting
## apply conditional formatting rules to excel cells
```csharp
//Create conditional formatting rule.
XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
xcfs1.AddRange(sheet.AllocatedRange);
IConditionalFormat format1 = xcfs1.AddCondition();
format1.FormatType = ConditionalFormatType.CellValue;
format1.FirstFormula = "800";
format1.Operator = ComparisonOperatorType.Greater;
format1.FontColor = Color.Red;
format1.BackColor = Color.LightSalmon;

//Create conditional formatting rule.
XlsConditionalFormats xcfs2 = sheet.ConditionalFormats.Add();
xcfs2.AddRange(sheet.AllocatedRange);
IConditionalFormat format2 = xcfs2.AddCondition();
format2.FormatType = ConditionalFormatType.CellValue;
format2.FirstFormula = "300";
format2.Operator = ComparisonOperatorType.Less;
format2.FontColor = Color.Green;
format2.BackColor = Color.LightBlue;
```

---

# spire.xls csharp conditional formatting
## apply data bars to cell range
```csharp
//Add data bars.
XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.AllocatedRange);
IConditionalFormat format = xcfs.AddCondition();
format.FormatType = ConditionalFormatType.DataBar;
format.DataBar.BarColor = Color.CadetBlue;
```

---

# spire.xls csharp gradient fill
## apply gradient fill effects to excel cells
```csharp
//Create a workbook
Workbook workbook = new Workbook();
workbook.Version = ExcelVersion.Version2010;

//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

//Get "B5" cell
CellRange range = sheet.Range["B5"];

//Set row height and column width
range.RowHeight = 50;
range.ColumnWidth = 30;
range.Text = "Hello";

//Set alignment style
range.Style.HorizontalAlignment = HorizontalAlignType.Center;

//Set gradient filling effects
range.Style.Interior.FillPattern = ExcelPatternType.Gradient;
range.Style.Interior.Gradient.ForeColor = Color.FromArgb(255, 255, 255);
range.Style.Interior.Gradient.BackColor = Color.FromArgb(79, 129, 189);
range.Style.Interior.Gradient.TwoColorGradient(GradientStyleType.Horizontal, GradientVariantsType.ShadingVariants1);
```

---

# spire.xls csharp conditional formatting
## apply icon sets to cell range
```csharp
//Add icon sets.
XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.AllocatedRange);
IConditionalFormat format = xcfs.AddCondition(); 
format.FormatType = ConditionalFormatType.IconSet;
format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1;
```

---

# spire.xls csharp colors and palette
## modify excel color palette and apply custom colors to cells
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Adding Orchid color to the palette at 60th index
workbook.ChangePaletteColor(Color.Orchid, 60);

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

CellRange cell = sheet.Range["B2"];
cell.Text = "Welcome to use Spire.XLS";

// Set the Orchid (custom) color to the font
cell.Style.Font.Color = Color.Orchid;
cell.Style.Font.Size = 20;
cell.AutoFitColumns();
cell.AutoFitRows();
```

---

# spire.xls csharp conditional formatting
## apply conditional formatting rules to excel cells at runtime
```csharp
private void AddComparisonRule1(Worksheet sheet)
{
    //Create conditional formatting rule
    XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
    xcfs1.AddRange(sheet.Range["A1:D1"]);
    IConditionalFormat cf1 = xcfs1.AddCondition();
    cf1.FormatType = ConditionalFormatType.CellValue;
    cf1.FirstFormula = "150";
    cf1.Operator = ComparisonOperatorType.Greater;
    cf1.FontColor = Color.Red;
    cf1.BackColor = Color.LightBlue;
}

private void AddComparisonRule2(Worksheet sheet)
{
    XlsConditionalFormats xcfs2 = sheet.ConditionalFormats.Add();
    xcfs2.AddRange(sheet.Range["A2:D2"]);
    IConditionalFormat cf2 = xcfs2.AddCondition();
    cf2.FormatType = ConditionalFormatType.CellValue;
    cf2.FirstFormula = "500";
    cf2.Operator = ComparisonOperatorType.Less;
    //Set border color
    cf2.LeftBorderColor = Color.Pink;
    cf2.RightBorderColor = Color.Pink;
    cf2.TopBorderColor = Color.DeepSkyBlue;
    cf2.BottomBorderColor = Color.DeepSkyBlue;
    cf2.LeftBorderStyle = LineStyleType.Medium;
    cf2.RightBorderStyle = LineStyleType.Thick;
    cf2.TopBorderStyle = LineStyleType.Double;
    cf2.BottomBorderStyle = LineStyleType.Double;
}

private void AddComparisonRule3(Worksheet sheet)
{
    //Create conditional formatting rule
    XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
    xcfs1.AddRange(sheet.Range["A3:D3"]);
    IConditionalFormat cf1 = xcfs1.AddCondition();
    cf1.FormatType = ConditionalFormatType.CellValue;
    cf1.FirstFormula = "300";
    cf1.SecondFormula = "500";
    cf1.Operator = ComparisonOperatorType.Between;
    cf1.BackColor = Color.Yellow;
}

private void AddComparisonRule4(Worksheet sheet)
{
    //Create conditional formatting rule
    XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
    xcfs1.AddRange(sheet.Range["A4:D4"]);
    IConditionalFormat cf1 = xcfs1.AddCondition();
    cf1.FormatType = ConditionalFormatType.CellValue;
    cf1.FirstFormula = "100";
    cf1.SecondFormula = "200";
    cf1.Operator = ComparisonOperatorType.NotBetween;
    //Set fill pattern type
    cf1.FillPattern = ExcelPatternType.ReverseDiagonalStripe;
    //Set foreground color
    cf1.Color = Color.FromArgb(255, 255, 0);
    //Set background color
    cf1.BackColor = Color.FromArgb(0, 255, 255);
}
```

---

# spire.xls csharp conditional formatting
## apply conditional formatting to dates in excel
```csharp
//Highlight cells that contain a date occurring in the last 7 days.
XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.AllocatedRange);
IConditionalFormat conditionalFormat = xcfs.AddTimePeriodCondition(TimePeriodType.Last7Days);
conditionalFormat.BackColor = Color.Orange;
```

---

# spire.xls csharp conditional formatting
## create formula-based conditional formatting in Excel
```csharp
//Set the conditional formatting formula and apply the rule to the chosen cell range.
XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(range);
IConditionalFormat conditional = xcfs.AddCondition();
conditional.FormatType = ConditionalFormatType.Formula;
conditional.FirstFormula = "=($A1<$B1)";
conditional.BackKnownColor = ExcelColors.Yellow;
```

---

# spire.xls csharp font styles
## apply various font styles to excel cells
```csharp
//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

//Set font style
sheet.Range["B1"].Style.Font.FontName = "Comic Sans MS";
sheet.Range["B2:D2"].Style.Font.FontName = "Corbel";
sheet.Range["B3:D7"].Style.Font.FontName = "Aleo";

//Set font size
sheet.Range["B1"].Style.Font.Size = 45;
sheet.Range["B2:D3"].Style.Font.Size = 25;
sheet.Range["B3:D7"].Style.Font.Size = 12;

//Set excel cell data to be bold
sheet.Range["B2:D2"].Style.Font.IsBold = true;

//Set excel cell data to be underline
sheet.Range["B3:B7"].Style.Font.Underline = FontUnderlineType.Single;

//set excel cell data color
sheet.Range["B1"].Style.Font.Color = Color.CornflowerBlue;
sheet.Range["B2:D2"].Style.Font.Color = Color.CadetBlue;
sheet.Range["B3:D7"].Style.Font.Color = Color.Firebrick;

//set excel cell data to be italic
sheet.Range["B3:D7"].Style.Font.IsItalic = true;

//Add strikethrough
sheet.Range["D3"].Style.Font.IsStrikethrough = true;
sheet.Range["D7"].Style.Font.IsStrikethrough = true;
```

---

# spire.xls csharp formatting
## set cell foreground and background colors and patterns
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

//Create a new style
CellStyle style = workbook.Styles.Add("newStyle1");

//Set filling pattern type
style.Interior.FillPattern = ExcelPatternType.VerticalStripe;

//Set filling Background color
style.Interior.Gradient.BackKnownColor = ExcelColors.Green;

//Set filling Foreground color
style.Interior.Gradient.ForeKnownColor = ExcelColors.Yellow;

//Apply the style to "B2" cell
sheet.Range["B2"].CellStyleName = style.Name;
sheet.Range["B2"].Text = "Test";
sheet.Range["B2"].RowHeight = 30;
sheet.Range["B2"].ColumnWidth = 50;

//Create a new style
style = workbook.Styles.Add("newStyle2");

//Set filling pattern type
style.Interior.FillPattern = ExcelPatternType.ThinHorizontalStripe;

//Set filling Foreground color
style.Interior.Gradient.ForeKnownColor = ExcelColors.Red;

//Apply the style to "B4" cell
sheet.Range["B4"].CellStyleName = style.Name;
sheet.Range["B4"].RowHeight = 30;
sheet.Range["B4"].ColumnWidth = 60;
```

---

# spire.xls csharp column formatting
## format a column in Excel with custom style
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

//Create a new style
CellStyle style = workbook.Styles.Add("newStyle");

//Set the vertical alignment of the text
style.VerticalAlignment = VerticalAlignType.Center;

//Set the horizontal alignment of the text
style.HorizontalAlignment = HorizontalAlignType.Center;

//Set the font color of the text
style.Font.Color = Color.Blue;

//Shrink the text to fit in the cell
style.ShrinkToFit = true;

//Set the bottom border color of the cell to OrangeRed
style.Borders[BordersLineType.EdgeBottom].Color = Color.OrangeRed;

//Set the bottom border type of the cell to Dotted
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Dotted;

//Apply the style to the first column
sheet.Columns[0].CellStyleName = style.Name;

sheet.Columns[0].Text = "Test";
```

---

# Excel Row Formatting
## Format a row in Excel with custom style properties
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

//Create a new style
CellStyle style = workbook.Styles.Add("newStyle");

//Set the vertical alignment of the text
style.VerticalAlignment = VerticalAlignType.Center;

//Set the horizontal alignment of the text
style.HorizontalAlignment = HorizontalAlignType.Center;

//Set the font color of the text
style.Font.Color = Color.Blue;

//Shrink the text to fit in the cell
style.ShrinkToFit = true;

//Set the bottom border color of the cell to OrangeRed
style.Borders[BordersLineType.EdgeBottom].Color = Color.OrangeRed;

//Set the bottom border type of the cell to Dotted
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Dotted;

//Apply the style to the second row
sheet.Rows[1].CellStyleName = style.Name;

sheet.Rows[1].Text = "Test";
```

---

# spire.xls csharp cell formatting
## create and apply cell style to range
```csharp
// Create a style
CellStyle style = workbook.Styles.Add("newStyle");
// Set the shading color
style.Color = Color.DarkGray;
// Set the font color
style.Font.Color = Color.White;
// Set font name
style.Font.FontName = "Times New Roman";
// Set font size
style.Font.Size = 12;
// Set bold for the font
style.Font.IsBold = true;
// Set text rotation
style.Rotation = 45;
// Set alignment
style.HorizontalAlignment = HorizontalAlignType.Center;
style.VerticalAlignment = VerticalAlignType.Center;

// Set the style for the specific range
workbook.Worksheets[0].Range["A1:J1"].CellStyleName = style.Name;
```

---

# spire.xls csharp conditional format
## get conditional format color from excel cells
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Load an existing Excel document
workbook.LoadFromFile("Template_Xls_13.xlsx");

// Get the first sheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Define a cell range
CellRange cRange = sheet.Range["A1:C1"];

// Retrieve the color of the condition format applied to the cell range
var color = cRange.GetConditionFormatsStyle().Color;

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp style
## get and set cell style
```csharp
//Get "B4" cell
CellRange range = sheet.Range["B4"];       
//Get the style of cell
CellStyle style = range.Style;
style.Font.FontName = "Calibri";
style.Font.IsBold = true;
style.Font.Size = 15;
style.Font.Color = Color.CornflowerBlue;

range.Style = style;
```

---

# Spire.XLS C# Conditional Formatting
## Highlight cells above and below average values
```csharp
// Add conditional format to highlight cells below average values
XlsConditionalFormats format1 = sheet.ConditionalFormats.Add();
// Set the cell range to apply the formatting
format1.AddRange(sheet.Range["E2:E10"]);
// Add below average condition
IConditionalFormat cf1 = format1.AddAverageCondition(AverageType.Below);
// Set background color for cells below average
cf1.BackColor = Color.SkyBlue; 

// Add conditional format to highlight cells above average values
XlsConditionalFormats format2 = sheet.ConditionalFormats.Add();
// Set the cell range to apply the formatting
format2.AddRange(sheet.Range["E2:E10"]);
// Add above average condition
IConditionalFormat cf2 = format2.AddAverageCondition(AverageType.Above);
// Set background color for cells above average
cf2.BackColor = Color.Orange;
```

---

# spire.xls csharp conditional formatting
## highlight duplicate and unique values in excel
```csharp
// Apply conditional formatting to highlight duplicate values in the range "C2:C10" with the color IndianRed.
XlsConditionalFormats duplicateFormats = sheet.ConditionalFormats.Add();
duplicateFormats.AddRange(sheet.Range["C2:C10"]);
IConditionalFormat duplicateCondition = duplicateFormats.AddCondition();
duplicateCondition.FormatType = ConditionalFormatType.DuplicateValues;
duplicateCondition.BackColor = Color.IndianRed;

// Apply conditional formatting to highlight unique values in the range "C2:C10" with the color Yellow.
XlsConditionalFormats uniqueFormats = sheet.ConditionalFormats.Add();
uniqueFormats.AddRange(sheet.Range["C2:C10"]);
IConditionalFormat uniqueCondition = uniqueFormats.AddCondition();
uniqueCondition.FormatType = ConditionalFormatType.UniqueValues;
uniqueCondition.BackColor = Color.Yellow;
```

---

# spire.xls csharp conditional formatting
## highlight top and bottom ranked values in excel ranges
```csharp
// Apply conditional formatting to range "D2:D10" to highlight the top 2 values
XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.Range["D2:D10"]);
IConditionalFormat format1 = xcfs.AddTopBottomCondition(TopBottomType.Top, 2);
format1.FormatType = ConditionalFormatType.TopBottom;
format1.BackColor = Color.Red;

// Apply conditional formatting to range "E2:E10" to highlight the bottom 2 values
XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
xcfs1.AddRange(sheet.Range["E2:E10"]);
IConditionalFormat format2 = xcfs1.AddTopBottomCondition(TopBottomType.Bottom, 2);
format2.FormatType = ConditionalFormatType.TopBottom;
format2.BackColor = Color.ForestGreen;
```

---

# spire.xls csharp indentation
## set cell text indentation level in Excel
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Add a new worksheet to the Excel object
Worksheet sheet = workbook.Worksheets[0];

//Access the "B5" cell from the worksheet
CellRange cell = sheet.Range["B5"];

//Add some value to the "B5" cell
cell.Text = "Hello Spire!";

//Set the indentation level of the text (inside the cell) to 2
cell.Style.IndentLevel = 2;
```

---

# spire.xls csharp interior formatting
## apply gradient fill and color to cell interiors
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Define the number of the colors
int maxColor = Enum.GetValues(typeof(ExcelColors)).Length;

// Create a random object
Random random = new Random(10000000);

for (int i = 2; i < 40; i++)
{
    // Random backKnownColor
    ExcelColors backKnownColor = (ExcelColors)(random.Next(1, maxColor / 2));

    // Add text
    sheet.Range["A1"].Text = "Color Name";
    sheet.Range["B1"].Text = "Red";
    sheet.Range["C1"].Text = "Green";
    sheet.Range["D1"].Text = "Blue";

    // Merge the sheet"E1-K1"
    sheet.Range["E1:K1"].Merge();
    sheet.Range["E1:K1"].Text = "Gradient";
    sheet.Range["A1:K1"].Style.Font.IsBold = true;
    sheet.Range["A1:K1"].Style.Font.Size = 11;

    // Set the text of color in sheetA-sheetD
    string colorName = backKnownColor.ToString();
    sheet.Range[string.Format("A{0}", i)].Text = colorName;
    sheet.Range[string.Format("B{0}", i)].NumberValue = workbook.GetPaletteColor(backKnownColor).R;
    sheet.Range[string.Format("C{0}", i)].NumberValue = workbook.GetPaletteColor(backKnownColor).G;
    sheet.Range[string.Format("D{0}", i)].NumberValue = workbook.GetPaletteColor(backKnownColor).B;

    // Merge the sheets
    sheet.Range[string.Format("E{0}:K{0}", i)].Merge();

    // Set the text of sheetE-sheetK
    sheet.Range[string.Format("E{0}:K{0}", i)].Text = colorName;

    //Set the interior of the color
    sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.FillPattern = ExcelPatternType.Gradient;
    sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.BackKnownColor = backKnownColor;
    sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.ForeKnownColor = ExcelColors.White;
    sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.GradientStyle = GradientStyleType.Vertical;
    sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.GradientVariant = GradientVariantsType.ShadingVariants1;
}

//AutoFit Column
sheet.AutoFitColumn(1);
```

---

# spire.xls csharp make cell active
## Set active cell and control visible area in Excel worksheet
```csharp
// Get the 2nd sheet
Worksheet sheet = workbook.Worksheets[1];

// Set the 2nd sheet as an active sheet
sheet.Activate();

// Set B2 cell as an active cell in the worksheet
sheet.SetActiveCell(sheet.Range["B2"]);

// Set the B column as the first visible column in the worksheet
sheet.FirstVisibleColumn = 1;

// Set the 2nd row as the first visible row in the worksheet
sheet.FirstVisibleRow = 1;
```

---

# Spire.XLS CSharp Number Formatting
## Demonstrate various number formatting styles in Excel cells
```csharp
// Input a number value for the specified cell and set the number format
sheet.Range["B10"].Text = "NUMBER FORMATTING";
sheet.Range["B10"].Style.Font.IsBold = true;

// Display as integer
sheet.Range["B13"].Text = "0"; 
sheet.Range["C13"].NumberValue = 1234.5678;
sheet.Range["C13"].NumberFormat = "0";

// Display as two decimal places
sheet.Range["B14"].Text = "0.00"; 
sheet.Range["C14"].NumberValue = 1234.5678;
sheet.Range["C14"].NumberFormat = "0.00";

// Display with thousand separator and two decimal places
sheet.Range["B15"].Text = "#,##0.00"; 
sheet.Range["C15"].NumberValue = 1234.5678;
sheet.Range["C15"].NumberFormat = "#,##0.00";

// Display as currency with thousand separator and two decimal places
sheet.Range["B16"].Text = "$#,##0.00"; 
sheet.Range["C16"].NumberValue = 1234.5678;
sheet.Range["C16"].NumberFormat = "$#,##0.00";

// Display positive numbers as is, negative numbers in red
sheet.Range["B17"].Text = "0;[Red]-0"; 
sheet.Range["C17"].NumberValue = -1234.5678;
sheet.Range["C17"].NumberFormat = "0;[Red]-0";

// Display positive numbers with two decimal places, negative numbers in red
sheet.Range["B18"].Text = "0.00;[Red]-0.00"; 
sheet.Range["C18"].NumberValue = -1234.5678;
sheet.Range["C18"].NumberFormat = "0.00;[Red]-0.00";

// Display positive numbers with thousand separator, negative numbers in red
sheet.Range["B19"].Text = "#,##0;[Red]-#,##0"; 
sheet.Range["C19"].NumberValue = -1234.5678;
sheet.Range["C19"].NumberFormat = "#,##0;[Red]-#,##0";

// Display positive numbers with thousand separator and two decimal places, negative numbers in red
sheet.Range["B20"].Text = "#,##0.00;[Red]-#,##0.000";
sheet.Range["C20"].NumberValue = -1234.5678;
sheet.Range["C20"].NumberFormat = "#,##0.00;[Red]-#,##0.00";

// Display as scientific notation with two decimal places
sheet.Range["B21"].Text = "0.00E+00"; 
sheet.Range["C21"].NumberValue = 1234.5678;
sheet.Range["C21"].NumberFormat = "0.00E+00";

// Display as percentage with two decimal places
sheet.Range["B22"].Text = "0.00%"; 
sheet.Range["C22"].NumberValue = 1234.5678;
sheet.Range["C22"].NumberFormat = "0.00%";

// Set background color for the range
sheet.Range["B13:B22"].Style.KnownColor = ExcelColors.Gray25Percent; 

// AutoFit Column
sheet.AutoFitColumn(2); 
sheet.AutoFitColumn(3);
```

---

# spire.xls csharp border formatting
## set cell borders in Excel worksheet
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Get the cell range where you want to apply border style
CellRange cr = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];

// Set the border style of the CellRange object to double line
cr.Borders.LineStyle = LineStyleType.Double;
// Set the diagonal down border style of the CellRange object to no line
cr.Borders[BordersLineType.DiagonalDown].LineStyle = LineStyleType.None;
// Set the diagonal up border style of the CellRange object to no line
cr.Borders[BordersLineType.DiagonalUp].LineStyle = LineStyleType.None;
// Set the border color of the CellRange object to CadetBlue
cr.Borders.Color = Color.CadetBlue;
```

---

# spire.xls csharp databar formatting
## set border to data bar in excel cells
```csharp
// Get the data bar format from the first conditional format
XlsConditionalFormats xcfs = sheet.ConditionalFormats[0];
IConditionalFormat cf = xcfs[0];
Spire.Xls.DataBar dataBar1 = cf.DataBar;

// Set the border type and color for the data bar format
dataBar1.BarBorder.Type = Spire.Xls.Core.Spreadsheet.ConditionalFormatting.DataBarBorderType.DataBarBorderSolid;
dataBar1.BarBorder.Color = Color.Red;

// Set a new data bar format to cell E1
sheet["E1"].NumberValue = 200;
XlsConditionalFormats xcfs2 = sheet.ConditionalFormats.Add();
xcfs2.AddRange(sheet.Range["E1"]);
IConditionalFormat cf2 = xcfs2.AddCondition();
cf2.FormatType = ConditionalFormatType.DataBar;
cf2.DataBar.BarBorder.Type = Spire.Xls.Core.Spreadsheet.ConditionalFormatting.DataBarBorderType.DataBarBorderSolid;
cf2.DataBar.BarBorder.Color = Color.Red;
cf2.DataBar.BarColor = Color.GreenYellow;
```

---

# Spire.XLS C# Conditional Formatting
## Set conditional format with formula in Excel
```csharp
// Add ConditionalFormat
XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();

// Define the range
xcfs.AddRange(sheet.Range["B5"]);

// Add condition
IConditionalFormat format = xcfs.AddCondition();
format.FormatType = ConditionalFormatType.CellValue;

// If greater than 1000
format.FirstFormula = "1000";
format.Operator = ComparisonOperatorType.Greater;
format.BackColor = Color.Orange;
```

---

# Spire.XLS C# Conditional Formatting
## Set row colors using conditional formatting
```csharp
//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Select the range that you want to format.
CellRange dataRange = sheet.AllocatedRange;

//Set conditional formatting.
XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(dataRange);
IConditionalFormat format1 = xcfs.AddCondition();
//Determines the cells to format.
format1.FirstFormula = "=MOD(ROW(),2)=0";
//Set conditional formatting type
format1.FormatType = ConditionalFormatType.Formula;
//Set the color.
format1.BackColor = Color.LightSeaGreen;

//Set the backcolor of the odd rows as Yellow.
XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
xcfs1.AddRange(dataRange);
IConditionalFormat format2 = xcfs.AddCondition(); 
format2.FirstFormula = "=MOD(ROW(),2)=1";
format2.FormatType = ConditionalFormatType.Formula;
format2.BackColor = Color.Yellow;
```

---

# spire.xls csharp conditional formatting
## set traffic lights icons in excel cells
```csharp
// Add a conditional formatting.
XlsConditionalFormats conditional = sheet.ConditionalFormats.Add();
conditional.AddRange(sheet.AllocatedRange);
IConditionalFormat format1 = conditional.AddCondition();

//Add a conditional formatting of cell range and set its type to CellValue.
format1.FormatType = ConditionalFormatType.CellValue;
format1.FirstFormula = "300";
format1.Operator = ComparisonOperatorType.Less;
format1.FontColor = Color.Black;
format1.BackColor = Color.LightSkyBlue;

//Add a conditional formatting of cell range and set its type to IconSet.
conditional.AddRange(sheet.AllocatedRange);
IConditionalFormat format = conditional.AddCondition();
format.FormatType = ConditionalFormatType.IconSet;
format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1;
```

---

# spire.xls csharp conditional formatting
## demonstrates how to apply various conditional formatting rules to Excel cells
```csharp
// Set row height
sheet.AllocatedRange.RowHeight = 15;
// Set column width
sheet.AllocatedRange.ColumnWidth = 16;

// Create conditional formatting rule for range A1:D1
XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
xcfs1.AddRange(sheet.Range["A1:D1"]);
IConditionalFormat cf1 = xcfs1.AddCondition();
cf1.FormatType = ConditionalFormatType.CellValue;
cf1.FirstFormula = "150";
cf1.Operator = ComparisonOperatorType.Greater;
cf1.FontColor = Color.Red;
cf1.BackColor = Color.LightBlue;

// Create conditional formatting rule for range A2:D2
XlsConditionalFormats xcfs2 = sheet.ConditionalFormats.Add();
xcfs2.AddRange(sheet.Range["A2:D2"]);
IConditionalFormat cf2 = xcfs2.AddCondition();
cf2.FormatType = ConditionalFormatType.CellValue;
cf2.FirstFormula = "300";
cf2.Operator = ComparisonOperatorType.Less;

//Set border color
cf2.LeftBorderColor = Color.Pink;
cf2.RightBorderColor = Color.Pink;
cf2.TopBorderColor = Color.DeepSkyBlue;
cf2.BottomBorderColor = Color.DeepSkyBlue;
cf2.LeftBorderStyle = LineStyleType.Medium;
cf2.RightBorderStyle = LineStyleType.Thick;
cf2.TopBorderStyle = LineStyleType.Double;
cf2.BottomBorderStyle = LineStyleType.Double;

//Add data bars
XlsConditionalFormats xcfs3 = sheet.ConditionalFormats.Add();
xcfs3.AddRange(sheet.Range["A3:D3"]);
IConditionalFormat cf3 = xcfs3.AddCondition();
cf3.FormatType = ConditionalFormatType.DataBar;
cf3.DataBar.BarColor = Color.CadetBlue;

//Add icon sets
XlsConditionalFormats xcfs4 = sheet.ConditionalFormats.Add();
xcfs4.AddRange(sheet.Range["A4:D4"]);
IConditionalFormat cf4 = xcfs4.AddCondition();
cf4.FormatType = ConditionalFormatType.IconSet;
cf4.IconSet.IconSetType = IconSetType.ThreeTrafficLights1;

//Add color scales
XlsConditionalFormats xcfs5 = sheet.ConditionalFormats.Add();
xcfs5.AddRange(sheet.Range["A5:D5"]);
IConditionalFormat cf5 = xcfs5.AddCondition();
cf5.FormatType = ConditionalFormatType.ColorScale;

//Highlight duplicate values in range "A6:D6" with BurlyWood color
XlsConditionalFormats xcfs6 = sheet.ConditionalFormats.Add();
xcfs6.AddRange(sheet.Range["A6:D6"]);
IConditionalFormat cf6 = xcfs6.AddCondition();
cf6.FormatType = ConditionalFormatType.DuplicateValues;
cf6.BackColor = Color.BurlyWood;
```

---

# spire.xls csharp text alignment
## set vertical and horizontal alignment and text rotation in excel cells
```csharp
// Set the vertical alignment to Top
sheet.Range["B1:C1"].Style.VerticalAlignment = VerticalAlignType.Top;

// Set the vertical alignment to Center
sheet.Range["B2:C2"].Style.VerticalAlignment = VerticalAlignType.Center;

// Set the vertical alignment of to Bottom
sheet.Range["B3:C3"].Style.VerticalAlignment = VerticalAlignType.Bottom;

// Set the horizontal alignment to General
sheet.Range["B4:C4"].Style.HorizontalAlignment = HorizontalAlignType.General;

// Set the horizontal alignment of to Left
sheet.Range["B5:C5"].Style.HorizontalAlignment = HorizontalAlignType.Left;

// Set the horizontal alignment of to Center
sheet.Range["B6:C6"].Style.HorizontalAlignment = HorizontalAlignType.Center;

// Set the horizontal alignment of to Right
sheet.Range["B7:C7"].Style.HorizontalAlignment = HorizontalAlignType.Right;

// Set the rotation degree
sheet.Range["B8:C8"].Style.Rotation = 45;

sheet.Range["B9:C9"].Style.Rotation = 90;

//Set the row height of cell
sheet.Range["B8:C9"].RowHeight = 60;
```

---

# spire.xls csharp text direction
## set text reading order in excel cell
```csharp
// Access the "B5" cell from the worksheet
CellRange cell = sheet.Range["B5"];

// Add some value to the "B5" cell
cell.Text = "Hello Spire!";

// Set the reading order from right to left of the text in the "B5" cell
cell.Style.ReadingOrder = ReadingOrderType.RightToLeft;
```

---

# Spire.XLS C# Style Application
## Apply custom predefined styles to Excel cells
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Create a new style
CellStyle style = workbook.Styles.Add("newStyle");
style.Font.FontName = "Calibri";
style.Font.IsBold = true;
style.Font.Size = 15;
style.Font.Color = Color.CornflowerBlue;

// Get the "B5" cell
CellRange range = sheet.Range["B5"];
range.Text = "Welcome to use Spire.XLS";

// Apply the newly created style to the cell
range.CellStyleName = style.Name;

// Autofit the columns for better display of cell content
range.AutoFitColumns();
```

---

# Spire.XLS C# Style Object
## Using style objects to format Excel cells
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Add a new worksheet to the Excel object
Worksheet sheet = workbook.Worksheets.Add("new sheet");

// Access the "B1" cell from the worksheet
CellRange cell = sheet.Range["B1"];

// Add some value to the "B1" cell
cell.Text = "Hello Spire!";

// Create a new style
CellStyle style = workbook.Styles.Add("newStyle");

// Set the vertical alignment of the text in the cell
style.VerticalAlignment = VerticalAlignType.Center;

// Set the horizontal alignment of the text in the cell
style.HorizontalAlignment = HorizontalAlignType.Center;

// Set the font color of the text in the cell
style.Font.Color = Color.Blue;

// Shrink the text to fit in the cell
style.ShrinkToFit = true;

// Set the bottom border color of the cell to GreenYellow
style.Borders[BordersLineType.EdgeBottom].Color = Color.GreenYellow;

// Set the bottom border type of the cell to Medium
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Medium;

// Assign the Style object to the cell
cell.Style = style;

// Apply the same style to some other cells
sheet.Range["B4"].Style = style;
sheet.Range["B4"].Text = "Test";
sheet.Range["C3"].CellStyleName = style.Name;
sheet.Range["C3"].Text = "Welcome to use Spire.XLS";
sheet.Range["D4"].Style = style;
```

---

# spire.xls csharp conditional formatting
## implement various conditional formatting types in Excel
```csharp
// Add IconSet conditional formatting type
private void AddDefaultIconSet(Worksheet sheet)
{
    XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
    xcfs.AddRange(sheet.Range["A1:C2"]);
    sheet.Range["A1:C2"].Style.FillPattern = ExcelPatternType.Solid;
    sheet.Range["A1:C2"].Style.Color = Color.Yellow;
    IConditionalFormat cf = xcfs.AddCondition();
    cf.FormatType = ConditionalFormatType.IconSet;
    sheet.Range["A1"].NumberValue = 0;
    sheet.Range["B1"].NumberValue = 3;
    sheet.Range["C1"].NumberValue = 6;
    sheet.Range["A2"].NumberValue = 2;
    sheet.Range["B2"].NumberValue = 5;
    sheet.Range["C2"].NumberValue = 8;
}

// Add ColorScale conditional formatting type
private void AddDefaultColorScale(Worksheet sheet)
{
    XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
    xcfs.AddRange(sheet.Range["A5:C6"]);
    sheet.Range["A5:C6"].Style.FillPattern = ExcelPatternType.Solid;
    sheet.Range["A5:C6"].Style.Color = Color.Pink;
    IConditionalFormat cf = xcfs.AddCondition();
    cf.FormatType = ConditionalFormatType.ColorScale;

    sheet.Range["A5"].NumberValue = 4;
    sheet.Range["B5"].NumberValue = 7;
    sheet.Range["C5"].NumberValue = 10;
    sheet.Range["A6"].NumberValue = 6;
    sheet.Range["B6"].NumberValue = 9;
    sheet.Range["C6"].NumberValue = 12;
}

// Add AboveAverage conditional formatting type
private void AddAboveAverage(Worksheet sheet)
{
    XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
    conds.AddRange(sheet.Range["A11:C12"]);
    sheet.Range["A11:C12"].Style.FillPattern = ExcelPatternType.Solid;
    sheet.Range["A11:C12"].Style.Color = Color.Tomato;
    IConditionalFormat cf = conds.AddAverageCondition(AverageType.Above);
    cf.FillPattern = ExcelPatternType.Solid;
    cf.BackColor = Color.Pink;

    sheet.Range["A11"].NumberValue = 10;
    sheet.Range["B11"].NumberValue = 13;
    sheet.Range["C11"].NumberValue = 16;
    sheet.Range["A12"].NumberValue = 12;
    sheet.Range["B12"].NumberValue = 15;
    sheet.Range["C12"].NumberValue = 18;
}

// Add Top10 conditional formatting type
private void AddTop10_1(Worksheet sheet)
{
    XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
    conds.AddRange(sheet.Range["A17:C20"]);
    sheet.Range["A17:C20"].Style.FillPattern = ExcelPatternType.Solid;
    sheet.Range["A17:C20"].Style.Color = Color.Gray;
    IConditionalFormat cf = conds.AddTopBottomCondition(TopBottomType.Top, 10);
    cf.FillPattern = ExcelPatternType.Solid;
    cf.BackColor = Color.Yellow;

    sheet.Range["A17"].NumberValue = 16;
    sheet.Range["B17"].NumberValue = 21;
    sheet.Range["C17"].NumberValue = 26;
    sheet.Range["A18"].NumberValue = 18;
    sheet.Range["B18"].NumberValue = 23;
    sheet.Range["C18"].NumberValue = 28;
    sheet.Range["A19"].NumberValue = 20;
    sheet.Range["B19"].NumberValue = 25;
    sheet.Range["C19"].NumberValue = 30;
    sheet.Range["A20"].NumberValue = 22;
    sheet.Range["B20"].NumberValue = 27;
    sheet.Range["C20"].NumberValue = 32;
}

// Add DataBars conditional formatting type
private void AddDataBar1(Worksheet sheet)
{
    // Add data bars
    XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
    xcfs.AddRange(sheet.Range["E1:G2"]);
    sheet.Range["E1:G2"].Style.FillPattern = ExcelPatternType.Solid;
    sheet.Range["E1:G2"].Style.Color = Color.YellowGreen;
    IConditionalFormat cf = xcfs.AddCondition();
    cf.FormatType = ConditionalFormatType.DataBar;
    cf.DataBar.BarColor = Color.Blue;
    cf.DataBar.MinPoint.Type = ConditionValueType.Percent;
    cf.DataBar.ShowValue = true;

    CellRange c = sheet.Range["E1"];
    c.NumberValue = 4;
    c = sheet.Range["F1"];
    c.NumberValue = 7;
    c = sheet.Range["G1"];
    c.NumberValue = 10;
    c = sheet.Range["E2"];
    c.NumberValue = 6;
    c = sheet.Range["F2"];
    c.NumberValue = 9;
    c = sheet.Range["G2"];
    c.NumberValue = 14;
}

// Add ContainsText conditional formatting type
private void AddContainsText(Worksheet sheet)
{
    XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
    conds.AddRange(sheet.Range["E5:G6"]);
    sheet.Range["E5:G6"].Style.FillPattern = ExcelPatternType.Solid;
    sheet.Range["E5:G6"].Style.Color = Color.LightBlue;
    IConditionalFormat cf = conds.AddContainsTextCondition("abc");
    cf.FillPattern = ExcelPatternType.Solid;
    cf.BackColor = Color.Yellow;

    CellRange c = sheet.Range["E5"];
    c.Text = "aa";
    c = sheet.Range["F5"];
    c.Text = "abfd";
    c = sheet.Range["G5"];
    c.Text = "aab";
    c = sheet.Range["E6"];
    c.Text = "abc";
    c = sheet.Range["F6"];
    c.Text = "cedf";
    c = sheet.Range["G6"];
    c.Text = "abcd";
}

// Add DuplicateValues conditional formatting type
private void AddDuplicate(Worksheet sheet)
{
    XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
    conds.AddRange(sheet.Range["E23:G24"]);
    sheet.Range["E23:G24"].Style.FillPattern = ExcelPatternType.Solid;
    sheet.Range["E23:G24"].Style.Color = Color.LightSlateGray;
    IConditionalFormat cf = conds.AddDuplicateValuesCondition();
    cf.FillPattern = ExcelPatternType.Solid;
    cf.BackColor = Color.Pink;
    CellRange c = sheet.Range["E23"];
    c.Text = "aa";
    c = sheet.Range["F23"];
    c.Text = "bb";
    c = sheet.Range["G23"];
    c.Text = "aa";
    c = sheet.Range["E24"];
    c.Text = "bbb";
    c = sheet.Range["F24"];
    c.Text = "bb";
    c = sheet.Range["G24"];
    c.Text = "ccc";
}

// Add TimePeriod conditional formatting type with Today attribute
private void AddTimePeriod_1(Worksheet sheet)
{
    XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
    conds.AddRange(sheet.Range["I1:K2"]);
    sheet.Range["I1:K2"].Style.FillPattern = ExcelPatternType.Solid;
    sheet.Range["I1:K2"].Style.Color = Color.LightSlateGray;
    IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.Today);
    cf.FillPattern = ExcelPatternType.Solid;
    cf.BackColor = Color.Pink;
    CellRange c = sheet.Range["I1"];
    c.Value2 = DateTime.Now.AddDays(-8).Date;

    c = sheet.Range["J1"];
    c.Value2 = DateTime.Now.AddDays(-7).Date;

    c = sheet.Range["K1"];
    c.Value2 = DateTime.Now.Date;

    c = sheet.Range["I2"];
    c.Text = "Today";

    c = sheet.Range["J2"];
    c.Value2 = DateTime.Now.AddDays(3).Date;

    c = sheet.Range["K2"];
    c.Value2 = DateTime.Now.AddMonths(2).Date;
}
```

---

# spire.xls csharp formula calculation
## calculate Excel formulas using Spire.XLS library
```csharp
// Create a workbook
Workbook workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

// Set column width
sheet.SetColumnWidth(1, 32);
sheet.SetColumnWidth(2, 16);
sheet.SetColumnWidth(3, 16);

// Set test data
sheet.Range[currentRow, 2].NumberValue = 7.3;
sheet.Range[currentRow, 3].NumberValue = 5;
sheet.Range[currentRow, 4].NumberValue = 8.2;
sheet.Range[currentRow, 5].NumberValue = 4;
sheet.Range[currentRow, 6].NumberValue = 3;
sheet.Range[currentRow, 7].NumberValue = 11.3;

// Set string formula
currentFormula = "=\"hello\"";
sheet.Range[++currentRow, 1].Text = "=\"hello\"";
sheet.Range[currentRow, 2].Formula = currentFormula;

// Set numeric formula
currentFormula = "=300";
sheet.Range[++currentRow, 1].Text = currentFormula;
sheet.Range[currentRow, 2].Formula = currentFormula;

// Set float formula
currentFormula = "=3389.639421";
sheet.Range[++currentRow, 1].Text = currentFormula;
sheet.Range[currentRow, 2].Formula = currentFormula;

// Set boolean formula
currentFormula = "=false";
sheet.Range[++currentRow, 1].Text = currentFormula;
sheet.Range[currentRow, 2].Formula = currentFormula;

// Set math operation formula
currentFormula = "=1+2+3+4+5-6-7+8-9";
sheet.Range[++currentRow, 1].Text = currentFormula;
sheet.Range[currentRow, 2].Formula = currentFormula;

// Set sheet reference formula
currentFormula = "=Sheet1!$B$3";
sheet.Range[++currentRow, 1].Text = currentFormula;
sheet.Range[currentRow, 2].Formula = currentFormula;

// Set function formula
currentFormula = "=AVERAGE(Sheet1!$D$3:G$3)";
sheet.Range[++currentRow, 1].Text = currentFormula;
sheet.Range[currentRow, 2].Formula = currentFormula;

// Set date function formula
currentFormula = "=NOW()";
sheet.Range[++currentRow, 1].Text = currentFormula;
sheet.Range[currentRow, 2].Formula = currentFormula;

// Set math function formula
currentFormula = "=SUM(18,29)";
sheet.Range[++currentRow, 1].Text = currentFormula;
sheet.Range[currentRow, 2].Formula = currentFormula;

// Calculate all formulas in workbook
workbook.CalculateAllValue();

// Calculate specific formula value
Object b3 = workbook.CalculateFormulaValue("Sheet1!$B$3");
Object c3 = workbook.CalculateFormulaValue("Sheet1!$C$3");
String formula = "Sheet1!$B$3 + Sheet1!$C$3";
Object value = workbook.CalculateFormulaValue(formula);

// Get formula and calculated value from cells
foreach (CellRange row in sheet["A5:B46"].Rows)
{
    String cellFormula = row.Columns[1].Formula;
    Object cellValue = row.Columns[1].FormulaValue;
    // Use formula and value as needed
}
```

---

# spire.xls csharp days formula
## Implement DAYS formula in Excel
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first sheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Add a formula to cell C4
sheet.Range["C4"].Formula = "=DAYS(A8,A1)";

// Calculate all values in the workbook
workbook.CalculateAllValue();
```

---

# spire.xls csharp named range formula
## insert formula with named range in excel
```csharp
// Create a workbook
Workbook workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

// Set values for cells A1 and A2
sheet.Range["A1"].Value = "1";
sheet.Range["A2"].Value = "1";

// Create a named range
INamedRange namedRange = workbook.NameRanges.Add("NewNamedRange");

// Set the local name and formula for the named range
namedRange.NameLocal = "=SUM(A1+A2)";

// Set the formula for cell C1 to reference the named range
sheet.Range["C1"].Formula = "NewNamedRange";
```

---

# Spire.XLS C# Formulas
## Demonstrating various Excel formulas implementation
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Write text values
sheet.Columns[0].NumberFormat = "@";
sheet.Range["A1"].Text = "=CEILING.MATH(-2.78, 5, -1)";
sheet.Range["A2"].Text = "=BITOR(23,10)";
sheet.Range["A3"].Text = "=BITAND(23,10)";
sheet.Range["A4"].Text = "=BITLSHIFT(23,2)";
sheet.Range["A5"].Text = "=BITRSHIFT(23,2)";
sheet.Range["A6"].Text = "=FLOOR.MATH(12.758, 2, -1)";
sheet.Range["A7"].Text = "=ISOWEEKNUM(DATE(2012, 1, 1))";
sheet.Range["A8"].Text = "=CEILING.PRECISE(-4.6, 3)";
sheet.Range["A9"].Text = "=ENCODEURL(\"https://www.e-iceblue.com\")";
sheet.Range["A10"].Text = "=ISFORMULA(A1)";
sheet.Range["A11"].Text = "=BITXOR(12, 58)";
// SPIREXLS-5395
sheet.Range["A12"].Text= "=BAHTTEXT(1234)";
//SPIREXLS-5393
sheet.Range["A13"].Text = "=TEXTBEFORE(\"Red riding hood's, red hood\", \"hood\")";
//SPIREXLS - 5394
sheet.Range["A14"].Text = "=TEXTSPLIT(A13,\" \", \".\", TRUE)";
//SPIREXLS-5397
sheet.Range["A15"].Text = "=TEXTAFTER(\"Red riding hood's, red hood\", \"hood\")";
//,SPIREXLS-5396
sheet.Range["A16"].Text = "= ARRAYTOTEXT(A1：B4，0)";
//SPIREXLS-5471
sheet.Range["A17"].Text = "=ARABIC(\"mcmxii\")";
//SPIREXLS-5478
sheet.Range["A18"].Text = "=BASE(15,2,10)";
//SPIREXLS-5479
sheet.Range["A19"].Text = "=COMBINA(3,10)";
//SPIREXLS-5480
sheet.Range["A20"].Text = "=XOR(3>12,2<9,4>6)";

// Write formulas
sheet.Range["B1"].Formula = "=CEILING.MATH(-2.78, 5, -1)";
sheet.Range["B2"].Formula = "=BITOR(23,10)";
sheet.Range["B3"].Formula = "=BITAND(23,10)";
sheet.Range["B4"].Formula = "=BITLSHIFT(23,2)";
sheet.Range["B5"].Formula = "=BITRSHIFT(23,2)";
sheet.Range["B6"].Formula = "=FLOOR.MATH(12.758, 2, -1)";
sheet.Range["B7"].Formula = "=ISOWEEKNUM(DATE(2012, 1, 1))";
sheet.Range["B8"].Formula = "=CEILING.PRECISE(-4.6, 3)";
sheet.Range["B9"].Formula = "=ENCODEURL(\"https://www.e-iceblue.com\")";
sheet.Range["B10"].Formula = "=ISFORMULA(A1)";
sheet.Range["B11"].Formula = "=BITXOR(12, 58)";
sheet.Range["B12"].Formula = "=BAHTTEXT(1234)";
sheet.Range["B13"].Formula = "=TEXTBEFORE(\"Red riding hood's, red hood\", \"hood\")";
sheet.Range["B14"].Formula = "=TEXTSPLIT(A13,\" \", \".\", TRUE)";
sheet.Range["B15"].Formula = "=TEXTAFTER(\"Red riding hood's, red hood\", \"hood\")";
sheet.Range["B16"].Formula = "=ARRAYTOTEXT(A1：B4，0)";
sheet.Range["B17"].Formula = "=ARABIC(\"mcmxii\")";
sheet.Range["B18"].Formula = "=BASE(15,2,10)";
sheet.Range["B19"].Formula = "=COMBINA(3,10)";
sheet.Range["B20"].Formula = "=XOR(3>12,2<9,4>6)";

// Calculate all value
workbook.CalculateAllValue();

// Autofit columns in the allocated range of the sheet
sheet.AllocatedRange.AutoFitColumns();
```

---

# spire.xls csharp read formulas
## Read Excel cell formulas and their calculated values
```csharp
// Create a workbook
Workbook workbook = new Workbook();
// Load an existing workbook from a file
workbook.LoadFromFile("ReadFormulas.xlsx");

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Get the formula from cell C14
string formula = sheet.Range["C14"].Formula;

// Get the numeric value resulting from the formula in cell C14
string formulaNumberValue = sheet.Range["C14"].FormulaNumberValue.ToString();
```

---

# spire.xls csharp addin function
## register and use add-in functions in excel
```csharp
String input = @"..\..\..\..\..\..\Data\Test.xlam";

// Create a workbook
Workbook workbook = new Workbook();

// Register AddIn function
workbook.AddInFunctions.Add(input, "TEST_UDF");
workbook.AddInFunctions.Add(input, "TEST_UDF1");

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Call AddIn function
sheet.Range["A1"].Formula = "=TEST_UDF()";
sheet.Range["A2"].Formula = "=TEST_UDF1()";
```

---

# spire.xls csharp formula handling
## remove formulas but keep calculated values
```csharp
// Loop through worksheets
foreach (Worksheet sheet in workbook.Worksheets)
{
    // Loop through cells
    foreach (CellRange cell in sheet.Range)
    {
        // If the cell contains formula, get the formula value, clear cell content, and then fill the formula value into the cell
        if (cell.HasFormula)
        {
            Object value = cell.FormulaValue;
            cell.Clear(ExcelClearOptions.ClearContent);
            cell.Value2 = value;
        }
    }
}
```

---

# spire.xls csharp subtotal formula
## demonstrates how to use SUBTOTAL formulas in Excel
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Set number values for cells A1:C3
sheet.Range["A1"].NumberValue = 1;
sheet.Range["A2"].NumberValue = 2;
sheet.Range["A3"].NumberValue = 3;
sheet.Range["B1"].NumberValue = 4;
sheet.Range["B2"].NumberValue = 5;
sheet.Range["B3"].NumberValue = 6;
sheet.Range["C1"].NumberValue = 7;
sheet.Range["C2"].NumberValue = 8;
sheet.Range["C3"].NumberValue = 9;

// Add SUBTOTAL formulas to calculate subtotal values
sheet.Range["A5"].Formula = "=SUBTOTAL(1,A1:C3)";
sheet.Range["B5"].Formula = "=SUBTOTAL(2,A1:C3)";
sheet.Range["C5"].Formula = "=SUBTOTAL(5,A1:C3)";

// Calculate all formulas in the workbook
workbook.CalculateAllValue();
```

---

# spire.xls csharp array formulas
## implement array formulas in Excel using Spire.XLS
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet =  workbook.Worksheets[0];

// Set number values for cells A1:C3
sheet.Range["A1"].NumberValue = 1;
sheet.Range["A2"].NumberValue = 2;
sheet.Range["A3"].NumberValue = 3;
sheet.Range["B1"].NumberValue = 4;
sheet.Range["B2"].NumberValue = 5;
sheet.Range["B3"].NumberValue = 6;
sheet.Range["C1"].NumberValue = 7;
sheet.Range["C2"].NumberValue = 8;
sheet.Range["C3"].NumberValue = 9;

// Write array formula
sheet.Range["A5:C6"].FormulaArray="=LINEST(A1:A3,B1:C3,TRUE,TRUE)";

// Calculate Formulas
workbook.CalculateAllValue();
```

---

# Spire.XLS C# Array R1C1 Formula
## Demonstrates how to use array formulas with R1C1 notation in Excel
```csharp
//Create a workbook and get the first sheet
Workbook workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

// Set number values for cells A1:C3
sheet.Range["A1"].NumberValue = 1;
sheet.Range["A2"].NumberValue = 2;
sheet.Range["A3"].NumberValue = 3;
sheet.Range["B1"].NumberValue = 4;
sheet.Range["B2"].NumberValue = 5;
sheet.Range["B3"].NumberValue = 6;
sheet.Range["C1"].NumberValue = 7;
sheet.Range["C2"].NumberValue = 8;
sheet.Range["C3"].NumberValue = 9;

// Set the text and alignment for cell B4
sheet.Range["B4"].Text = "Sum:";
sheet.Range["B4"].Style.HorizontalAlignment = HorizontalAlignType.Right;

// Write an array R1C1 formula in cell C4 to calculate the sum
sheet.Range["C4"].FormulaArrayR1C1 = "=SUM(R[-3]C[-2]:R[-1]C)";

// Calculate all formulas in the workbook
workbook.CalculateAllValue();
```

---

# spire.xls csharp r1c1 formula
## demonstrate how to use R1C1-style formulas in Excel
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Set number values for cells A1:C3
sheet.Range["A1"].NumberValue = 1;
sheet.Range["A2"].NumberValue = 2;
sheet.Range["A3"].NumberValue = 3;
sheet.Range["B1"].NumberValue = 4;
sheet.Range["B2"].NumberValue = 5;
sheet.Range["B3"].NumberValue = 6;
sheet.Range["C1"].NumberValue = 7;
sheet.Range["C2"].NumberValue = 8;
sheet.Range["C3"].NumberValue = 9;

// Set the text and alignment for cell B4
sheet.Range["B4"].Text = "Sum:";
sheet.Range["B4"].Style.HorizontalAlignment = HorizontalAlignType.Right;

// Write R1C1 formula
sheet.Range["C4"].FormulaR1C1 = "=SUM(R[-3]C[-2]:R[-1]C)";

// Calculate all formulas in the workbook
workbook.CalculateAllValue();
```

---

# spire.xls csharp formulas
## write various types of formulas to Excel cells
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Set column widths
sheet.SetColumnWidth(1, 32);
sheet.SetColumnWidth(2, 16);

// Add header
sheet.Range[1, 1].Value = "Examples of formulas :";
sheet.Range[2, 1].Value = "Test data:";

// Apply formatting to the header cell
CellRange range = sheet.Range["A1"];
range.Style.Font.IsBold = true;
range.Style.FillPattern = ExcelPatternType.Solid;
range.Style.KnownColor = ExcelColors.LightGreen1;

// Test data
sheet.Range[2, 2].NumberValue = 7.3;
sheet.Range[2, 3].NumberValue = 5;
sheet.Range[2, 4].NumberValue = 8.2;

// Formula headers
sheet.Range[3, 1].Value = "Formulas";
sheet.Range[3, 2].Value = "Results";
range = sheet.Range[3, 1, 3, 2];
range.Style.Font.IsBold = true;
range.Style.KnownColor = ExcelColors.LightGreen1;

// String formula
string currentFormula = "=\"hello\"";
sheet.Range[4, 1].NumberFormat = "@";
sheet.Range[4, 1].Text = "=\"hello\"";
sheet.Range[4, 2].Formula = currentFormula;

// Numeric formula
currentFormula = "=300";
sheet.Range[5, 1].NumberFormat = "@";
sheet.Range[5, 1].Text = currentFormula;
sheet.Range[5, 2].Formula = currentFormula;

// Boolean formula
currentFormula = "=false";
sheet.Range[6, 1].NumberFormat = "@";
sheet.Range[6, 1].Text = currentFormula;
sheet.Range[6, 2].Formula = currentFormula;

// Mathematical expression
currentFormula = "=1+2+3+4+5-6-7+8-9";
sheet.Range[7, 1].NumberFormat = "@";
sheet.Range[7, 1].Text = currentFormula;
sheet.Range[7, 2].Formula = currentFormula;

// Cell reference
currentFormula = "=Sheet1!$B$3";
sheet.Range[8, 1].NumberFormat = "@";
sheet.Range[8, 1].Text = currentFormula;
sheet.Range[8, 2].Formula = currentFormula;

// Function with range reference
currentFormula = "=AVERAGE(Sheet1!$D$3:G$3)";
sheet.Range[9, 1].NumberFormat = "@";
sheet.Range[9, 1].Text = currentFormula;
sheet.Range[9, 2].Formula = currentFormula;

// Statistical function
currentFormula = "=Count(3,5,8,10,2,34)";
sheet.Range[10, 1].NumberFormat = "@";
sheet.Range[10, 1].Text = currentFormula;
sheet.Range[10, 2].Formula = currentFormula;

// Date function
currentFormula = "=NOW()";
sheet.Range[11, 1].NumberFormat = "@";
sheet.Range[11, 1].Text = currentFormula;
sheet.Range[11, 2].Formula = currentFormula;
sheet.Range[11, 2].Style.NumberFormat = "yyyy-MM-DD";

// Mathematical function
currentFormula = "=SQRT(40)";
sheet.Range[12, 1].NumberFormat = "@";
sheet.Range[12, 1].Text = currentFormula;
sheet.Range[12, 2].Formula = currentFormula;

// Logical function
currentFormula = "=IF(4,2,2)";
sheet.Range[13, 1].NumberFormat = "@";
sheet.Range[13, 1].Text = currentFormula;
sheet.Range[13, 2].Formula = currentFormula;
```

---

# Spire.XLS C# Header Footer
## Add image to first page header and footer
```csharp
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

sheet.PageSetup.DifferentFirst = (byte)1;

// Load an image from disk
Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");

// Set the image header
sheet.PageSetup.FirstLeftHeaderImage = image;
sheet.PageSetup.FirstCenterHeaderImage = image;
sheet.PageSetup.FirstRightHeaderImage = image;

// Set the image footer
sheet.PageSetup.FirstLeftFooterImage = image;
sheet.PageSetup.FirstCenterFooterImage = image;
sheet.PageSetup.FirstRightFooterImage = image;

// Set the view mode of the sheet
sheet.ViewMode = ViewMode.Layout;
```

---

# spire.xls csharp watermark
## add watermark to excel worksheets
```csharp
// Create a font 
Font font = new System.Drawing.Font("Arial", 40);
String watermark = "Confidential";

foreach (Worksheet sheet in workbook.Worksheets)
{
    // Call DrawText() to create an image
    System.Drawing.Image imgWtrmrk = DrawText(watermark, font, System.Drawing.Color.LightCoral, System.Drawing.Color.White, sheet.PageSetup.PageHeight, sheet.PageSetup.PageWidth);

    // Set image as left header image
    sheet.PageSetup.LeftHeaderImage = imgWtrmrk;
    sheet.PageSetup.LeftHeader = "&G";

    // The watermark will only appear in this mode, it will disappear if the mode is normal
    sheet.ViewMode = ViewMode.Layout;
}

private static System.Drawing.Image DrawText(String text, System.Drawing.Font font, Color textColor, Color backColor, double height, double width)
{
    // Create a bitmap image with specified width and height
    Image img = new Bitmap((int)width, (int)height);
    Graphics drawing = Graphics.FromImage(img);

    // Get the size of text
    SizeF textSize = drawing.MeasureString(text, font);

    // Set rotation point
    drawing.TranslateTransform(((int)width - textSize.Width) / 2, ((int)height - textSize.Height) / 2);

    // Rotate text
    drawing.RotateTransform(-45);

    //  Reset translate transform    
    drawing.TranslateTransform(-((int)width - textSize.Width) / 2, -((int)height - textSize.Height) / 2);

    // Paint the background
    drawing.Clear(backColor);

    // Create a brush for the text
    Brush textBrush = new SolidBrush(textColor);

    // Draw text on the image at center position
    drawing.DrawString(text, font, textBrush, ((int)width - textSize.Width) / 2, ((int)height - textSize.Height) / 2);
    drawing.Save();
    return img;
}
```

---

# Spire.XLS C# Header Footer
## Change font and size for header and footer in Excel
```csharp
// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Set the new font and size for the header and footer
string text = sheet.PageSetup.LeftHeader;

// "Arial Unicode MS" is font name, "18" is font size
text = "&\"Arial Unicode MS\"&18 Header Footer Sample by Spire.XLS ";
sheet.PageSetup.LeftHeader = text;
sheet.PageSetup.RightFooter = text;
```

---

# Spire.XLS C# Header Footer
## Set different headers and footers for odd and even pages
```csharp
// Get the first worksheet
Worksheet sheet = wb.Worksheets[0];

// Set text for the range
sheet.Range["A1"].Text = "Page 1";
sheet.Range["G1"].Text = "Page 2";

// Set the different header footer for Odd and Even pages
sheet.PageSetup.DifferentOddEven = 1;

// Set the header with font, size, bold and color
sheet.PageSetup.OddHeaderString = "&\"Arial\"&12&B&KFFC000 Odd_Header";
sheet.PageSetup.OddFooterString = "&\"Arial\"&12&B&KFFC000 Odd_Footer";
sheet.PageSetup.EvenHeaderString = "&\"Arial\"&12&B&KFF0000 Even_Header";
sheet.PageSetup.EvenFooterString = "&\"Arial\"&12&B&KFF0000 Even_Footer";

// Set view mode as page layout view
sheet.ViewMode = ViewMode.Layout;
```

---

# spire.xls csharp header footer
## set different header footer on first page
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Set the value to show the headers/footers for first page are different from the other pages.
sheet.PageSetup.DifferentFirst = 1;
     
// Set the header and footer for the first page.
sheet.PageSetup.FirstHeaderString = "Different First page";
sheet.PageSetup.FirstFooterString = "Different First footer";

// Set the other pages' header and footer. 
sheet.PageSetup.LeftHeader = "Demo of Spire.XLS";
sheet.PageSetup.CenterFooter = "Footer by Spire.XLS";
```

---

# spire.xls csharp header footer
## set image header and footer in excel worksheet
```csharp
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Load an image from disk
Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");

// Set the image header
sheet.PageSetup.LeftHeaderImage = image;
sheet.PageSetup.LeftHeader = "&G";

// Set the image footer
sheet.PageSetup.CenterFooterImage = image;
sheet.PageSetup.CenterFooter = "&G";

// Set the view mode of the sheet
sheet.ViewMode = ViewMode.Layout;
```

---

# spire.xls csharp header footer
## set crop position for images in header and footer
```csharp
// Set the cropping values for the left header picture
sheet.PageSetup.LeftHeaderPictureCropTop = 0.2f;
sheet.PageSetup.LeftHeaderPictureCropBottom = 0.3f;
sheet.PageSetup.LeftHeaderPictureCropLeft = 0.3f;
sheet.PageSetup.LeftHeaderPictureCropRight = 0.2f;

// Set the cropping values for the left footer picture
sheet.PageSetup.LeftFooterPictureCropTop = 0.2f;
sheet.PageSetup.LeftFooterPictureCropBottom = 0.3f;
sheet.PageSetup.LeftFooterPictureCropLeft = 0.3f;
sheet.PageSetup.LeftFooterPictureCropRight = 0.2f;

// Set the cropping values for the center header picture
sheet.PageSetup.CenterHeaderPictureCropTop = 0.3f;
sheet.PageSetup.CenterHeaderPictureCropBottom = 0.4f;
sheet.PageSetup.CenterHeaderPictureCropLeft = 0.4f;
sheet.PageSetup.CenterHeaderPictureCropRight = 0.3f;

// Set the cropping values for the center footer picture
sheet.PageSetup.CenterFooterPictureCropTop = 0.3f;
sheet.PageSetup.CenterFooterPictureCropBottom = 0.4f;
sheet.PageSetup.CenterFooterPictureCropLeft = 0.4f;
sheet.PageSetup.CenterFooterPictureCropRight = 0.3f;

// Set the cropping values for the right header picture
sheet.PageSetup.RightHeaderPictureCropTop = 0.2f;
sheet.PageSetup.RightHeaderPictureCropBottom = 0.3f;
sheet.PageSetup.RightHeaderPictureCropLeft = 0.9f;
sheet.PageSetup.RightHeaderPictureCropRight = 0.4f;

// Set the cropping values for the right footer picture
sheet.PageSetup.RightFooterPictureCropTop = 0.2f;
sheet.PageSetup.RightFooterPictureCropBottom = 0.3f;
sheet.PageSetup.RightFooterPictureCropLeft = 0.9f;
sheet.PageSetup.RightFooterPictureCropRight = 0.4f;
```

---

# spire.xls csharp header footer
## set Excel header and footer
```csharp
// Set left header,"Arial Unicode MS" is font name, "18" is font size.
Worksheet.PageSetup.LeftHeader = "&\"Arial Unicode MS\"&14 Spire.XLS for .NET ";

// Set center footer 
Worksheet.PageSetup.CenterFooter = "Footer Text";

// Set view mode as  page layout view
Worksheet.ViewMode = ViewMode.Layout;
```

---

# spire.xls csharp hyperlink
## add hyperlink to text in excel cells
```csharp
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Add url link
HyperLink UrlLink = sheet.HyperLinks.Add(sheet.Range["D10"]);
// Set display text
UrlLink.TextToDisplay = sheet.Range["D10"].Text;
// Set url link type
UrlLink.Type = HyperLinkType.Url;
// Set url address
UrlLink.Address = "http://en.wikipedia.org/wiki/Chicago";

//Add email link
HyperLink MailLink = sheet.HyperLinks.Add(sheet.Range["E10"]);
// Set display text
MailLink.TextToDisplay = sheet.Range["E10"].Text;
// Set mail link type
MailLink.Type = HyperLinkType.Url;
// Set mail address
MailLink.Address = "mailto:Amor.Aqua@gmail.com";
```

---

# spire.xls csharp image hyperlink
## add image hyperlink to excel worksheet
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load a Workbook from disk
Worksheet sheet = workbook.Worksheets[0];

// Set width for the first column
sheet.Columns[0].ColumnWidth = 22;
// Set value for cell "A1"
sheet.Range["A1"].Text = "Image Hyperlink";
// Set vertical alignment as top
sheet.Range["A1"].Style.VerticalAlignment = VerticalAlignType.Top;

// Insert an image to a specific cell
ExcelPicture picture = sheet.Pictures.Add(2, 1, imagePath);

// Add a hyperlink to the image
picture.SetHyperLink("https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html", true);
```

---

# Spire.XLS C# Get Hyperlink Types
## Extract hyperlink addresses and types from an Excel worksheet
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Iterate all hyperlinks
foreach (var item in sheet.HyperLinks)
{
    // Get hyperlink address
    string address = item.Address;
    // Get hyperlink type
    HyperLinkType type = item.Type;
}
```

---

# spire.xls csharp get image hyperlink
## Extract hyperlink address from an image in Excel
```csharp
//Create a workbook
Workbook workbook = new Workbook();

//Get the first picture of the first worksheet
ExcelPicture picture = workbook.Worksheets[0].Pictures[0];

//Get the address
string address = picture.GetHyperLink().Address;
```

---

# spire.xls csharp hyperlink
## create hyperlink to external file in excel
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Get cell "A1"
CellRange range = sheet.Range[1, 1];

// Add a hyperlink within the specified range
HyperLink hyperlink = sheet.HyperLinks.Add(range);

// Set the hyperlink type
hyperlink.Type = HyperLinkType.File;

// Set the display text for the hyperlink
hyperlink.TextToDisplay = "Link To External File";

// Set the file address for the hyperlink
hyperlink.Address = "..\\..\\..\\..\\..\\..\\Data\\SampleB_4.xlsx";
```

---

# spire.xls csharp hyperlink
## create hyperlink to another sheet cell
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
  
// Get cell "A1"
CellRange range = sheet.Range["A1"];

// Add hyperlink in the range
HyperLink hyperlink = sheet.HyperLinks.Add(range);

// Set the link type
hyperlink.Type = HyperLinkType.Workbook;

// Set the display text
hyperlink.TextToDisplay = "Link to Sheet2 cell C5";

// Set the link address
hyperlink.Address = "Sheet2!C5";
```

---

# spire.xls csharp hyperlink modification
## modify hyperlink text and address in excel worksheet
```csharp
// Get the collection of all hyperlinks in the worksheet
HyperLinksCollection links = sheet.HyperLinks;

// Modify the values of TextToDisplay and Address properties of the first hyperlink
links[0].TextToDisplay = "Spire.XLS for .NET";
links[0].Address = "http://www.e-iceblue.com/Introduce/excel-for-net-introduce.html";
```

---

# spire.xls csharp read hyperlinks
## read hyperlink addresses from excel worksheet
```csharp
// Load an existing workbook
Workbook workbook = new Workbook();
workbook.LoadFromFile("ReadHyperlinks.xlsx");

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Read hyperlink addresses
string firstHyperlink = sheet.HyperLinks[0].Address;
string secondHyperlink = sheet.HyperLinks[1].Address;

// Dispose of the workbook
workbook.Dispose();
```

---

# spire.xls csharp hyperlinks
## remove hyperlinks from excel worksheet
```csharp
//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Get the collection of all hyperlinks in the worksheet
HyperLinksCollection links = sheet.HyperLinks;

// Remove the content of cells B1, B2, B3 to remove link text
sheet.Range["B1"].ClearAll();
sheet.Range["B2"].ClearAll();
sheet.Range["B3"].ClearAll();

// Remove hyperlink
sheet.HyperLinks.RemoveAt(0);
sheet.HyperLinks.RemoveAt(0);
sheet.HyperLinks.RemoveAt(0);
```

---

# spire.xls csharp hyperlinks
## retrieve external file hyperlinks from excel worksheet
```csharp
// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

StringBuilder content = new StringBuilder();

// Retrieve external file hyperlinks.
foreach (HyperLink item in sheet.HyperLinks)
{
    String address = item.Address;
    String sheetName = item.Range.WorksheetName;
    CellRange range = item.Range;
    content.AppendLine(String.Format("Cell[{0},{1}] in sheet \"" + sheetName + "\" contains File URL: {2}", range.Row, range.Column, address));
}
```

---

# spire.xls csharp hyperlinks
## write hyperlinks to excel cells
```csharp
// Set the text for cell B9 as "Home page"
sheet.Range["B9"].Text = "Home page";

// Add a hyperlink to cell B10
HyperLink hylink1 = sheet.HyperLinks.Add(sheet.Range["B10"]);
hylink1.Type = HyperLinkType.Url;
hylink1.Address = @"http://www.e-iceblue.com";

// Set the text for cell B11 as "Support"
sheet.Range["B11"].Text = "Support";

// Add a hyperlink to cell B12
HyperLink hylink2 = sheet.HyperLinks.Add(sheet.Range["B12"]);
hylink2.Type = HyperLinkType.Url;
hylink2.Address = "mailto:support@e-iceblue.com";

// Set the text for cell B13 as "Forum"
sheet.Range["B13"].Text = "Forum";

// Add a hyperlink to cell B14
HyperLink hylink3 = sheet.HyperLinks.Add(sheet.Range["B14"]);
hylink3.Type = HyperLinkType.Url;
hylink3.Address = "https://www.e-iceblue.com/forum/";
```

---

# spire.xls csharp custom object
## Add custom objects to Excel using MarkerDesigner
```csharp
public class Student
{
    internal Student(string name, int age)
    {
        this.Name = name;
        this.Age = age;
    }
    public string Name { get; set; }
    public int Age { get; set; }
}

// Create a new workbook
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Set marker designer field in cells
sheet.Range["A1"].Value = "&=Student.Name";
sheet.Range["B1"].Value = "&=Student.Age";

// Create a list of custom objects
List<Student> list = new List<Student>();
list.Add(new Student("John", 16));
list.Add(new Student("Mary", 17));
list.Add(new Student("Lucy", 17));

// Fill custom object using the "Student" parameter
workbook.MarkerDesigner.AddParameter("Student", list);
workbook.MarkerDesigner.Apply();
workbook.CalculateAllValue();

// AutoFit rows and columns to adjust their sizes based on content
sheet.AllocatedRange.AutoFitRows();
sheet.AllocatedRange.AutoFitColumns();
```

---

# spire.xls csharp marker designer
## add variable array to excel using marker designer
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Set marker designer field in cell A1
sheet.Range["A1"].Value = "&=Array";

// Fill an array using the "Array" parameter
workbook.MarkerDesigner.AddArray("Array", new string[] { "Spire.Xls", "Spire.Doc", "Spire.PDF", "Spire.Presentation", "Spire.Email" });
workbook.MarkerDesigner.Apply();
workbook.CalculateAllValue();

// AutoFit rows and columns to adjust their sizes based on content
sheet.AllocatedRange.AutoFitRows();
sheet.AllocatedRange.AutoFitColumns();
```

---

# spire.xls csharp copy cell style
## copy cell style using Marker Designer
```csharp
// Create a DataTable
DataTable dt = new DataTable("data");

// Define columns in the DataTable
dt.Columns.Add(new DataColumn("name", typeof(string)));
dt.Columns.Add(new DataColumn("age", typeof(int)));

// Add three rows to the DataTable
DataRow drName1 = dt.NewRow();
DataRow drName2 = dt.NewRow();
DataRow drName3 = dt.NewRow();

drName1["name"] = "John";
drName1["age"] = 15;
drName2["name"] = "Jess";
drName2["age"] = 22;
drName3["name"] = "Alan";
drName3["age"] = 36;

dt.Rows.Add(drName1);
dt.Rows.Add(drName2);
dt.Rows.Add(drName3);

// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Fill the DataTable using the "data" parameter in Marker Designer
workbook.MarkerDesigner.AddDataTable("data", dt);
workbook.MarkerDesigner.Apply();
```

---

# spire.xls marker designer detect blank
## Detect blank cells using Marker Designer in Excel
```csharp
// Create a DataSet
DataSet ds = new DataSet();

// Fill the DataSet from an XML file
ds.ReadXml(@"..\..\..\..\..\..\Data\Data.xml");

// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Fill a DataTable using the "data" parameter in Marker Designer
workbook.MarkerDesigner.AddDataTable("data", ds.Tables["data"]);
workbook.MarkerDesigner.Apply();

// Calculate all formulas in the workbook
workbook.CalculateAllValue();
```

---

# spire.xls csharp marker designer
## apply parameters and datatable to excel template using marker designer
```csharp
// Fill a parameter named "Variable1" with the value 1234.5678
workbook.MarkerDesigner.AddParameter("Variable1", 1234.5678);

// Fill a DataTable named "Country" with the data from dt
workbook.MarkerDesigner.AddDataTable("Country", dt);
workbook.MarkerDesigner.Apply();

// AutoFit rows and columns to adjust their sizes based on content
sheet.AllocatedRange.AutoFitRows();
sheet.AllocatedRange.AutoFitColumns();
```

---

# spire.xls marker designer
## set data direction using marker designer
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Load an existing workbook from a file
workbook.LoadFromFile("MarkerDesigner2.xlsx");

// Create a DataTable named "data"
DataTable dt = new DataTable("data");

// Define a column named "value" in the DataTable
dt.Columns.Add(new DataColumn("value", typeof(string)));

// Create three new rows for the DataTable
DataRow drName1 = dt.NewRow();
DataRow drName2 = dt.NewRow();
DataRow drName3 = dt.NewRow();

// Set values for the "value" column in each row
drName1["value"] = "Text1";
drName2["value"] = "Text2";
drName3["value"] = "Text3";

// Add the rows to the DataTable
dt.Rows.Add(drName1);
dt.Rows.Add(drName2);
dt.Rows.Add(drName3);

// Add the DataTable to the Marker Designer with the parameter name "data"
workbook.MarkerDesigner.AddDataTable("data", dt);

// Apply the changes made in the Marker Designer
workbook.MarkerDesigner.Apply();

// Save the modified workbook to the specified file using Excel 2013 format
workbook.SaveToFile("SetDataDirection_result.xlsx", ExcelVersion.Version2013);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS C# Named Range Formatting
## Format cells within a named range in Excel
```csharp
// Get specific named range by index
INamedRange NamedRange = workbook.NameRanges[0];

// Get the cell range of the named range
IXLSRange range = NamedRange.RefersToRange;

// Set color for the range
range.Style.Color = Color.Yellow;

// Set the font as bold
range.Style.Font.IsBold = true;
```

---

# Spire.XLS C# Named Ranges
## Retrieve all named ranges from an Excel workbook
```csharp
// Get all named ranges in the workbook
INameRanges ranges = workbook.NameRanges;

// Iterate over each named range
foreach (INamedRange nameRange in ranges)
{
    // Process the name of the current named range
    string rangeName = nameRange.Name;
}
```

---

# spire.xls csharp named range
## get address of named range in excel
```csharp
// Create a new workbook and load an existing document from a file
Workbook workbook = new Workbook();
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AllNamedRanges.xlsx");

// Get a specific named range by its index
INamedRange NamedRange = workbook.NameRanges[0];

// Get the address of the named range
string address = NamedRange.RefersToRange.RangeAddress;
```

---

# spire.xls csharp named range
## get named range of cell range
```csharp
// Determine whether NamedRange exists in Range A7:D7
var result = workbook.Worksheets[0].Range["A7:D7"].GetNamedRange();

// Determine whether NamedRange exists in Range A4:D4
var result1 = workbook.Worksheets[0].Range["A4:D4"].GetNamedRange();

// Determine whether NamedRange exists in cell C14
var result2 = workbook.Worksheets[0].Range["C14"].GetNamedRange();
if (result2 == null)
{
    // C14 cell does not have NameRange
}
```

---

# Spire.XLS C# Named Range Operations
## Get specific named ranges by index and name
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load an existing workbook from a file
workbook.LoadFromFile("AllNamedRanges.xlsx");

// Get a specific named range by its index
string name1 = workbook.NameRanges[1].Name;

// Get a specific named range by its name
string name2 = workbook.NameRanges["NameRange3"].Name;

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp named ranges
## merge named range cells in excel
```csharp
// Get a specific named range by its index
INamedRange NamedRange = workbook.NameRanges[0];

// Get the range of the named range
IXLSRange range = NamedRange.RefersToRange;

// Merge cells within the range
range.Merge();
```

---

# spire.xls csharp named ranges
## create and configure named range in excel workbook
```csharp
// Create a new named range
INamedRange NamedRange = workbook.NameRanges.Add("NewNamedRange");

// Set the range of the named range to cover cells A8 to E12 on the worksheet
NamedRange.RefersToRange = sheet.Range["A8:E12"];
```

---

# spire.xls csharp named range
## remove named ranges from Excel workbook
```csharp
// Remove the named range by its index
workbook.NameRanges.RemoveAt(0);

// Remove the named range by its name
workbook.NameRanges.Remove("NameRange2");
```

---

# Spire.XLS C# Rename Named Range
## Rename a named range in an Excel workbook
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Load an existing workbook from a file
workbook.LoadFromFile("AllNamedRanges.xlsx");

// Rename the named range at index 0 to "RenameRange"
workbook.NameRanges[0].Name = "RenameRange";

// Save the modified workbook
workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010);
```

---

# Spire.XLS C# Named Range Formula
## Set formula using named range in Excel
```csharp
// Create a named range
INamedRange NamedRange = workbook.NameRanges.Add("MyNamedRange");

// Refers to range
NamedRange.RefersToRange = sheet.Range["B10:B12"];

//Set the formula of range to named range
sheet.Range["B13"].Formula = "=SUM(MyNamedRange)";
```

---

# spire.xls csharp extract msg ole object
## extract MSG OLE objects from Excel worksheet
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

OleObjectType type;
// Determine if there is an ole object in the sheet
if (sheet.HasOleObjects)
{
    for (int i = 0; i < sheet.OleObjects.Count; i++)
    {
        var Object = sheet.OleObjects[i];
        // Get the type of ole object
        type = sheet.OleObjects[i].ObjectType;
        switch (type)
        {
            // If the type of ole object is msg
            case OleObjectType.Msg:
                File.WriteAllBytes(outputFile, Object.OleData);
                break;
        }
    }
}
```

---

# spire.xls csharp ole objects
## extract OLE objects from Excel worksheet
```csharp
// Check if the worksheet contains any OLE objects
if (sheet.HasOleObjects)
{
    // Iterate over each OLE object in the worksheet
    for (int i = 0; i < sheet.OleObjects.Count; i++)
    {
        // Get the current OLE object
        var oleObject = sheet.OleObjects[i];

        // Determine the type of the OLE object
        OleObjectType type = oleObject.ObjectType;

        // Perform operations based on the type of the OLE object
        switch (type)
        {
            // Word document
            case OleObjectType.WordDocument:
                File.WriteAllBytes("Ole.docx", oleObject.OleData);
                break;
            // PowerPoint document
            case OleObjectType.PowerPointSlide:
                File.WriteAllBytes("Ole.pptx", oleObject.OleData);
                break;
            // PDF document
            case OleObjectType.AdobeAcrobatDocument:
                File.WriteAllBytes("Ole.pdf", oleObject.OleData);
                break;
        }
    }
}
```

---

# Spire.XLS C# OLE Objects
## Get origin name and type of OLE objects in Excel worksheet
```csharp
// Get the first sheet
Worksheet worksheet = workbook.Worksheets[0];

String information = "";
// Check if the worksheet contains any OLE objects
if (worksheet.HasOleObjects)
{
    // Iterate over each OLE object in the worksheet
    for (int i = 0; i < worksheet.OleObjects.Count; i++)
    {
        // Get the current OLE object
        var Object = worksheet.OleObjects[i];

        // Determine the type of the OLE object
        OleObjectType type = worksheet.OleObjects[i].ObjectType;
        information += "Type: " + type.ToString()+"\n";

        // Determine the origin name of the OLE object
        String originName = worksheet.OleObjects[i].OleOriginName;
        information += "Origin Name: " + originName + "\n";
    }
}
```

---

# Spire.XLS OLE Objects
## Insert OLE objects into Excel worksheet
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet ws = workbook.Worksheets[0];

// Set the text in cell A1
ws.Range["A1"].Text = "Here is an OLE Object.";

// Insert an OLE object
string xlsFile = @"..\..\..\..\..\..\Data\InsertOLEObjects.xls";
Image image = GenerateImage(xlsFile);
IOleObject oleObject = ws.OleObjects.Add(xlsFile, image, OleLinkType.Embed);
oleObject.Location = ws.Range["B4"];
oleObject.ObjectType = OleObjectType.ExcelWorksheet;

// Generate image for OLE object representation
private Image GenerateImage(string fileName)
{
    Workbook book = new Workbook();
    book.LoadFromFile(fileName);
    // Set page margins to zero
    book.Worksheets[0].PageSetup.LeftMargin = 0;
    book.Worksheets[0].PageSetup.RightMargin = 0;
    book.Worksheets[0].PageSetup.TopMargin = 0;
    book.Worksheets[0].PageSetup.BottomMargin = 0;
    // Convert worksheet range to image
    return book.Worksheets[0].ToImage(1, 1, 19, 5);
}
```

---

# spire.xls csharp oleobject
## insert WAV file as OLE object in Excel worksheet
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Add an OLE object
IOleObject oleObject = sheet.OleObjects.Add(@"..\..\..\..\..\..\Data\WAVFileSample.wav", Image.FromFile(@"..\..\..\..\..\..\Data\SpireXls.png"), OleLinkType.Embed);

// Set the location for the OLE object
oleObject.Location = sheet.Range["B4"];

// Set the type of the OLE object as a package
oleObject.ObjectType = OleObjectType.Package;
```

---

# spire.xls csharp page setup
## get Excel paper dimensions
```csharp
// Create a new workbook
Workbook workbook = new Workbook();

// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Get the dimensions of A2 paper
sheet.PageSetup.PaperSize = PaperSizeType.A2Paper;
float a2Width = sheet.PageSetup.PageWidth;
float a2Height = sheet.PageSetup.PageHeight;

// Get the dimensions of A3 paper
sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
float a3Width = sheet.PageSetup.PageWidth;
float a3Height = sheet.PageSetup.PageHeight;

// Get the dimensions of A4 paper
sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
float a4Width = sheet.PageSetup.PageWidth;
float a4Height = sheet.PageSetup.PageHeight;

// Get the dimensions of letter-sized paper
sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
float letterWidth = sheet.PageSetup.PageWidth;
float letterHeight = sheet.PageSetup.PageHeight;

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp page setup
## set Excel page order type
```csharp
// Get the reference of the PageSetup of the worksheet
PageSetup pageSetup = sheet.PageSetup;

// Set the order type of the pages to "Over then down"
pageSetup.Order = OrderType.OverThenDown;
```

---

# spire.xls csharp page setup
## set Excel paper size to A4
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Set the paper size of the worksheet as A4 paper.
sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
```

---

# Spire.XLS C# Page Setup
## Set first page number of Excel worksheet
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Set the first page number of the worksheet pages.
sheet.PageSetup.FirstPageNumber = 2;
```

---

# spire.xls csharp page setup
## set header and footer margins in Excel worksheet
```csharp
// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Get the PageSetup object of the first worksheet.
PageSetup pageSetup = sheet.PageSetup;

// Set the margins of header and footer.
pageSetup.HeaderMarginInch = 2;
pageSetup.FooterMarginInch = 2;
```

---

# spire.xls csharp page setup
## set margins of excel sheet
```csharp
// Get the PageSetup object of the worksheet.
PageSetup pageSetup = sheet.PageSetup;

// Set bottom,left,right and top page margins.
pageSetup.BottomMargin = 2;
pageSetup.LeftMargin = 1;
pageSetup.RightMargin = 1;
pageSetup.TopMargin = 3;
```

---

# Spire.XLS C# Print Setup
## Configure worksheet printing options
```csharp
// Get the reference of the PageSetup of the worksheet.
PageSetup pageSetup = sheet.PageSetup;

// Allow to print gridlines.
pageSetup.IsPrintGridlines = true;

// Allow to print row/column headings.
pageSetup.IsPrintHeadings = true;

// Allow to print worksheet in black & white mode.
pageSetup.BlackAndWhite = true;

// Allow to print comments as displayed on worksheet.
pageSetup.PrintComments = PrintCommentType.InPlace;

// Allow to print worksheet with draft quality.
pageSetup.Draft = true;

// Allow to print cell errors as N/A.
pageSetup.PrintErrors = PrintErrorsType.NA;
```

---

# spire.xls csharp page setup
## set page orientation to landscape
```csharp
// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Set the page orientation to Landscape. 
sheet.PageSetup.Orientation = PageOrientationType.Landscape;
```

---

# spire.xls csharp print area
## set print area for excel worksheet
```csharp
// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Get the reference of the PageSetup of the worksheet.
PageSetup pageSetup = sheet.PageSetup;

// Specify the cells range of the print area.
pageSetup.PrintArea = "A1:E5";
```

---

# spire.xls csharp print quality
## set print quality of excel file
```csharp
//Set the print quality of the worksheet to 180 dpi.
sheet.PageSetup.PrintQuality = 180;
```

---

# spire.xls csharp print setup
## set print title rows and columns in excel worksheet
```csharp
// Get the PageSetup of the worksheet
PageSetup pageSetup = sheet.PageSetup;

// Define columns A and B as title columns
pageSetup.PrintTitleColumns = "$A:$B";

// Define rows 1 and 2 as title rows
pageSetup.PrintTitleRows = "$1:$2";
```

---

# spire.xls csharp page setup
## set worksheet fit to page properties
```csharp
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Set the FitToPagesTall property to 1
sheet.PageSetup.FitToPagesTall = 1;

// Set the FitToPagesWide property to 1
sheet.PageSetup.FitToPagesWide = 1;
```

---

# spire.xls csharp page setup
## center worksheet on page
```csharp
// Get the PageSetup object of the first page
PageSetup pageSetup = sheet.PageSetup;

// Set the worksheet to be centered horizontally on the page
pageSetup.CenterHorizontally = true;

// Set the worksheet to be centered vertically on the page
pageSetup.CenterVertically = true;
```

---

# spire.xls csharp pivot table filters
## add filters to pivot table fields
```csharp
// Retrieve the first pivot table from the second sheet
XlsPivotTable pt = workbook.Worksheets[1].PivotTables[0] as XlsPivotTable;

// Add a label filter to the first row field of the pivot table
pt.RowFields[0].AddLabelFilter(PivotLabelFilterType.Between, "Argentina", "Nicaragua");

// Add a value filter on the first row field of the pivot table
pt.RowFields[0].AddValueFilter(PivotValueFilterType.LessThan, pt.DataFields[0], 5300000, null);

// Calculate the pivot table data after applying filters
pt.CalculateData();
```

---

# Spire.XLS C# Pivot Table
## Change pivot table data source
```csharp
// Define the range of cells to be used as the new data source
CellRange range = sheet.Range["A1:C15"];

// Get the first pivot table from the second worksheet
PivotTable table = workbook.Worksheets[1].PivotTables[0] as PivotTable;

// Change the data source of the pivot table to the new range
table.ChangeDataSource(range);

// Disable automatic refresh of the pivot table cache on load
table.Cache.IsRefreshOnLoad = false;
```

---

# Spire.XLS C# Clear Pivot Fields
## Demonstrates how to clear data fields in a pivot table using Spire.XLS library
```csharp
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];

// Get the first pivot table from the sheet
XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

// Clear all the data fields in the pivot table
pt.DataFields.Clear();

// Calculate the pivot table data
pt.CalculateData();
```

---

# spire.xls csharp pivot table consolidation
## apply consolidation functions to pivot table data fields
```csharp
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];

XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

// Apply Average consolidation function to first data field
pt.DataFields[0].Subtotal = SubtotalTypes.Average;

// Apply Max consolidation function to second data field
pt.DataFields[1].Subtotal = SubtotalTypes.Max;

// Calculate data
pt.CalculateData();
```

---

# spire.xls csharp pivottable
## create pivot table in excel
```csharp
// Add a PivotTable to the worksheet
CellRange dataRange = sheet.Range["A1:C7"];
PivotCache cache = workbook.PivotCaches.Add(dataRange);
PivotTable pt = sheet.PivotTables.Add("Pivot Table", sheet.Range["E10"], cache);

// Drag the fields to the row area.
PivotField pf = pt.PivotFields["Product"] as PivotField;
pf.Axis = AxisTypes.Row;
PivotField pf2 = pt.PivotFields["Month"] as PivotField;
pf2.Axis = AxisTypes.Row;

// Drag the field to the data area.
pt.DataFields.Add(pt.PivotFields["Count"], "SUM of Count", SubtotalTypes.Sum);

// Set PivotTable style
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12;

// Autofit columns generated by the pivotTable
pt.CalculateData();
sheet.AutoFitColumn(5);
sheet.AutoFitColumn(6);
```

---

# spire.xls csharp pivot table
## customize pivot table field names
```csharp
// Access the first pivot table in the worksheet
XlsPivotTable pivotTable = sheet.PivotTables[0] as XlsPivotTable;

// Set a custom name for the row field
pivotTable.RowFields[0].CustomName = "custom_rowName";

// Set a custom name for the column field
pivotTable.ColumnFields[0].CustomName = "custom_colName";

// Set a custom name for the data field
pivotTable.DataFields[0].CustomName = "custom_DataName";

// Calculate the pivot table data
pivotTable.CalculateData();
```

---

# Spire.XLS C# Pivot Table
## Disable Pivot Table Ribbon
```csharp
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];

// Get the first pivot table from the sheet 
XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

//Disable ribbon for this pivot table
pt.EnableWizard = false;
```

---

# spire.xls csharp pivot table
## expand or collapse rows in pivot table
```csharp
// Get the first pivot table from the sheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTable = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;

// Calculate data
pivotTable.CalculateData();

// Collapse the rows
(pivotTable.PivotFields["Vendor No"] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3501", true);

// Expand the rows
(pivotTable.PivotFields["Vendor No"] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3502", false);
```

---

# spire.xls csharp pivot table
## format pivot table data field
```csharp
// Get the first pivot table from the sheet
XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;
// Access the data field.
PivotDataField pivotDataField = pt.DataFields[0];

// Set data display format
pivotDataField.ShowDataAs = PivotFieldFormatType.PercentageOfColumn;
```

---

# spire.xls csharp pivot table formatting
## format pivot table appearance
```csharp
// Get the first pivot table from the worksheet
XlsPivotTable pivotTable = sheet.PivotTables[0] as XlsPivotTable;

// Set the built-in style for the pivot table appearance
pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleLight10;

// Enable the display of grid drop zone in the pivot table
pivotTable.Options.ShowGridDropZone = true;

// Set the row layout type to compact in the pivot table
pivotTable.Options.RowLayout = PivotTableLayoutType.Compact;
```

---

# spire.xls csharp pivot table
## get pivot table refresh information
```csharp
// Get first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];

// Get the first pivot table
XlsPivotTable pivotTable = worksheet.PivotTables[0] as XlsPivotTable;

// Get the refreshed information
DateTime dateTime = pivotTable.Cache.RefreshDate;
string refreshedBy = pivotTable.Cache.RefreshedBy;

// Set string format for displaying
string result = string.Format("Pivot table refreshed by:  " + refreshedBy + "\r\nPivot table refreshed date: " + dateTime.ToString());
```

---

# spire.xls csharp pivot table
## group pivot table by date
```csharp
// Get the first pivot table in the worksheet
XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

// Get the first row field in the pivot table
IPivotField field = pt.RowFields[0];

// Set the start and end dates for grouping
DateTime start = new DateTime(2023, 1, 5);
DateTime end = new DateTime(2023, 3, 2);

// Set the group by type to days
PivotGroupByTypes[] types = new PivotGroupByTypes[] { PivotGroupByTypes.Days };

// Create a new group with the specified start and end dates, group by type, and interval
field.CreateGroup(start, end, types, 10);

// Calculate the pivot table data
pt.CalculateData();

// Refresh the pivot table cache
pt.Cache.IsRefreshOnLoad = true;
```

---

# spire.xls csharp pivot table layout
## set pivot table layout to tabular format
```csharp
// Get the first worksheet
Worksheet worksheet = workbook.Worksheets[0];

// Get the first PivotTable
XlsPivotTable xlsPivotTable = (XlsPivotTable)worksheet.PivotTables[0];

// Set the PivotTable layout type
xlsPivotTable.Options.ReportLayout = PivotTableLayoutType.Tabular;
```

---

# spire.xls csharp pivottable
## refresh pivot table data
```csharp
// Update the data source of PivotTable.
sheet.Range["D2"].Value = "999";

// Get the PivotTable that was built on the data source.
XlsPivotTable pt = workbook.Worksheets[0].PivotTables[0] as XlsPivotTable;

// Refresh the data of PivotTable.
pt.Cache.IsRefreshOnLoad = true;
```

---

# spire.xls csharp pivot table
## repeat all item labels for pivot table
```csharp
// Iterate through each pivot table in the "Pivot" worksheet
foreach (XlsPivotTable pt in workbook.Worksheets["Pivot"].PivotTables)
{
    // Set the RepeatAllItemLabels property to true for the pivot table
    pt.Options.RepeatAllItemLabels = true;

    // Calculate the data for the pivot table
    pt.CalculateData();

    // Refresh the cache for the pivot table
    pt.Cache.IsRefreshOnLoad = true;
}
```

---

# Spire.XLS C# Pivot Table
## Enable Repeat Item Labels in Pivot Table
```csharp
// Create a pivot cache using the data range
PivotCache cache = workbook.PivotCaches.Add(dataRange);

// Add a pivot table to the pivot sheet using the pivot cache
PivotTable pt = sheet2.PivotTables.Add("Pivot Table", sheet.Range["A1"], cache);

// Set the VendorNo field as a row field and specify its header caption
var r1 = pt.PivotFields["VendorNo"];
r1.Axis = AxisTypes.Row;
pt.Options.RowHeaderCaption = "VendorNo";
r1.Subtotals = SubtotalTypes.None;

// Enable repeating item labels for the VendorNo field
r1.RepeatItemLabels = true;

// Enable repeating item labels for the OnHand field
pt.PivotFields["OnHand"].RepeatItemLabels = true;

// Set the row layout type to tabular
pt.Options.RowLayout = PivotTableLayoutType.Tabular;

// Set the Desc field as an additional row field
var r2 = pt.PivotFields["Desc"];
r2.Axis = AxisTypes.Row;

// Add the OnHand field as a data field with the label "Sum of onHand"
pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.None);

// Set the built-in style for the pivot table appearance
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12;
```

---

# spire.xls csharp pivot table format options
## set format options for pivot table
```csharp
// Get the sheet where the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];

// Access the first pivot table in the sheet
XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

// Enable automatic formatting for the pivot table report
pt.Options.IsAutoFormat = true;

// Show grand totals for rows in the pivot table report
pt.ShowRowGrand = true;

// Show grand totals for columns in the pivot table report
pt.ShowColumnGrand = true;

// Display a custom string in cells that contain null values
pt.DisplayNullString = true;
pt.NullString = "null";

// Set the layout of the pivot table report
pt.PageFieldOrder = PagesOrderType.DownThenOver;
```

---

# Spire.XLS C# Pivot Table Field Formatting
## Set format options for pivot table fields including sort type, subtotal display, and auto show
```csharp
// Access the first pivot table in the worksheet
XlsPivotTable pivotTable = sheet.PivotTables[0] as XlsPivotTable;

// Access the first pivot field in the pivot table
PivotField pivotField = pivotTable.PivotFields[0] as PivotField;

// Set the sort type of the pivot field to ascending
pivotField.SortType = PivotFieldSortType.Ascending;

// Enable displaying subtotals at the top of groups for the pivot field
pivotField.SubtotalTop = true;

// Set the subtotal type of the pivot field to Count
pivotField.Subtotals = SubtotalTypes.Count;

// Enable auto show for the pivot field
pivotField.IsAutoShow = true;
```

---

# Spire.XLS C# Pivot Table Conditional Formatting
## Set conditional formatting for pivot table fields
```csharp
// Get the worksheet with the PivotTable
Worksheet worksheet = workbook.Worksheets["PivotTable"];

// Get the PivotTable from the worksheet
PivotTable table = (PivotTable)worksheet.PivotTables[0];

// Add a conditional format to the PivotTable
PivotConditionalFormatCollection pcfs = table.PivotConditionalFormats;
PivotConditionalFormat pc = pcfs.AddPivotConditionalFormat(table.DataFields[0]);
Spire.Xls.Core.IConditionalFormat cf = pc.AddCondition();
cf.FormatType = ConditionalFormatType.NotContainsBlanks;
cf.FillPattern = ExcelPatternType.Solid;
cf.BackColor = Color.Yellow;
```

---

# spire.xls csharp pivot table
## show data field in row area of pivot table
```csharp
// Access the pivot table in the worksheet
XlsPivotTable pivotTable = sheet.PivotTables[0] as XlsPivotTable;

// Show the data field in the row area of the pivot table
pivotTable.ShowDataFieldInRow = true;

// Calculate the data in the pivot table
pivotTable.CalculateData();
```

---

# Spire.XLS C# Pivot Table Subtotals
## Enable subtotals display in Excel pivot table
```csharp
// Get the worksheet that contains the pivot table
Worksheet sheet = workbook.Worksheets["Pivot Table"];

// Get the first pivot table from the worksheet
XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

// Enable the display of subtotals in the pivot table
pt.ShowSubtotals = true;
```

---

# Spire.XLS C# Pivot Table Sorting
## Create and sort a pivot table in Excel using Spire.XLS library
```csharp
// Add an empty worksheet to the workbook and set its name
Worksheet sheet2 = workbook.CreateEmptySheet();
sheet2.Name = "Pivot Table";

// Specify the data source range for the pivot table
CellRange dataRange = sheet.Range["A1:C9"];

// Create a pivot cache using the data range
PivotCache cache = workbook.PivotCaches.Add(dataRange);

// Add a pivot table to the second worksheet using the specified cache
PivotTable pt = sheet2.PivotTables.Add("Pivot Table", sheet.Range["A1"], cache);

// Configure the pivot table settings
PivotField r1 = pt.PivotFields["No"] as PivotField;
r1.Axis = AxisTypes.Row;
pt.Options.RowLayout = PivotTableLayoutType.Tabular;

// Sort the "No" field in descending order
r1.SortType = PivotFieldSortType.Descending;

PivotField r2 = pt.PivotFields["Name"] as PivotField;
r2.Axis = AxisTypes.Row;

// Add a data field to the pivot table
pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.None);

// Set the pivot table style
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12;
```

---

# spire.xls csharp pivot table
## update pivot table data source and refresh
```csharp
// Access the "Data" worksheet
Worksheet data = workbook.Worksheets["Data"];

// Modify the data source by changing the value in cell A2 to "NewValue"
data.Range["A2"].Text = "NewValue";

// Modify the data source by changing the value in cell D2 to 28000
data.Range["D2"].NumberValue = 28000;

// Access the worksheet containing the pivot table
Worksheet sheet = workbook.Worksheets["PivotTable"];

// Get the first pivot table from the worksheet
XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

// Set the pivot table's cache to refresh on load
pt.Cache.IsRefreshOnLoad = true;

// Calculate and update the pivot table data
pt.CalculateData();
```

---

# Spire.XLS Custom Paper Size for Printing
## Set custom paper size and print Excel document
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load file from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CustomPaperSizeForPrinting.xlsx");

// Get the first worksheet 
Worksheet worksheet = workbook.Worksheets[0];

// Set the paper size to the printer's custom paper size
worksheet.PageSetup.CustomPaperSizeName = "customPaper";

//Custom the paper size directly
//sheet.PageSetup.SetCustomPaperSize(224, (float)50);

//Set the page orientation
//sheet.PageSetup.Orientation = PageOrientationType.Portrait;

// Use the default printer to print
workbook.PrintDocument.Print();

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS C# Gray Level Print
## Print Excel document with gray level settings
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Set the GrayLevelForPrint to true
workbook.ConverterSetting.GrayLevelForPrint = true;

// Print this document
workbook.PrintDocument.Print();

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel Page Setup for Printing
## Configure print settings for Excel worksheets
```csharp
// Specifying the print area
PageSetup pageSetup = worksheet.PageSetup;
pageSetup.PrintArea = "A1:E19";

// Define column A & E as title columns
pageSetup.PrintTitleColumns = "$A:$E";

// Define row numbers 1 as title rows
pageSetup.PrintTitleRows = "$1:$2";

// Allow to print with gridlines
pageSetup.IsPrintGridlines = true;

// Allow to print with row/column headings
pageSetup.IsPrintHeadings = true;

// Allow to print worksheet in black & white mode
pageSetup.BlackAndWhite = true;

// Allow to print comments as displayed on worksheet
pageSetup.PrintComments = PrintCommentType.InPlace;

// Set printing quality
pageSetup.PrintQuality = 150;

// Allow to print cell errors as N/A
pageSetup.PrintErrors = PrintErrorsType.NA;

// Set the printing order 
pageSetup.Order = OrderType.OverThenDown;

// Print file
workbook.PrintDocument.Print();
```

---

# spire.xls csharp print
## Print Excel document
```csharp
// Create a workbook
Workbook workbook = new Workbook();

//Load the Excel document from disk
workbook.LoadFromFile("PrintExcel.xlsx");

// Access the printer settings of the workbook's print document
PrinterSettings settings = workbook.PrintDocument.PrinterSettings;

// Specify the range of pages to be printed (from page 0 to page 1)
settings.FromPage = 0;
settings.ToPage = 1;

// Use the default printer to print
workbook.PrintDocument.Print();

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp digital signature
## add digital signature to excel workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Add a digital certificate for signing the workbook
String inputFile_pfx = @"..\..\..\..\..\..\Data\gary.pfx";            
X509Certificate2 cert = new X509Certificate2(inputFile_pfx, "e-iceblue");

// Specify the date and time for the digital signature
DateTime certtime = new DateTime(2020, 7, 1, 7, 10, 36);

// Add a digital signature to the workbook using the provided certificate, signer name ("e-iceblue"), and signature timestamp
IDigitalSignatures dsc = workbook.AddDigitalSignature(cert, "e-iceblue", certtime);
```

---

# spire.xls csharp add signature line
## add a signature line to an Excel worksheet
```csharp
//Create a workbook instance
Workbook workbook = new Workbook();

//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Add a signature line 
sheet.Range["A1"].AddSignatureLine("Rose","manager", "manager@test.com", "a short text" ,false,true);
```

---

# spire.xls csharp security
## detect if excel workbook is password protected
```csharp
//Detect if the Excel workbook is password protected
bool value = Workbook.IsPasswordProtected(input);

if (value)
{
    textBox1.Text = "Yes";
}
else
{
    textBox1.Text = "No";
}
```

---

# spire.xls csharp security
## hide formulas in excel worksheet and protect with password
```csharp
// Hide the formulas in the used range
sheet.AllocatedRange.IsFormulaHidden = true;

// Protect the worksheet with password
sheet.Protect("e-iceblue");
```

---

# spire.xls csharp security
## lock specific cells in excel worksheet
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Create an empty worksheet.
workbook.CreateEmptySheet();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Loop through all the rows in the worksheet and unlock them.
for (int i = 0; i < 255; i++)
{
    sheet.Rows[i].Style.Locked = false;
}

// Lock specific cell in the worksheet.
sheet.Range["A1"].Text = "Locked";
sheet.Range["A1"].Style.Locked = true;

// Lock specific cell range in the worksheet.
sheet.Range["C1:E3"].Text = "Locked";
sheet.Range["C1:E3"].Style.Locked = true;

// Set the password.
sheet.Protect("123", SheetProtectionType.All);
```

---

# Spire.XLS Security
## Lock specific column in new Excel file
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Create an empty worksheet.
workbook.CreateEmptySheet();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Loop through all the columns in the worksheet and unlock them.
for (int i = 0; i < 255; i++)
{
    sheet.Rows[i].Style.Locked = false;
}

// Lock the fourth column in the worksheet.
sheet.Columns[3].Text = "Locked";
sheet.Columns[3].Style.Locked = true;

// Set the password.
sheet.Protect("123", SheetProtectionType.All);
```

---

# spire.xls csharp security
## lock specific row in excel worksheet
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Create an empty worksheet.
workbook.CreateEmptySheet();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Loop through all the rows in the worksheet and unlock them.
for (int i = 0; i < 255; i++)
{
    sheet.Rows[i].Style.Locked = false;
}

// Lock the third row in the worksheet.
sheet.Rows[2].Text = "Locked";
sheet.Rows[2].Style.Locked = true;

// Set the password.
sheet.Protect("123", SheetProtectionType.All);
```

---

# spire.xls csharp cell protection
## protect specific cells in excel worksheet
```csharp
// Protect cell
sheet.Range["B3"].Style.Locked = true;
sheet.Range["C3"].Style.Locked = false;

// Set password
sheet.Protect("TestPassword", SheetProtectionType.All);
```

---

# Spire.XLS C# Worksheet Protection
## Protect worksheet with editable ranges
```csharp
// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Define the specified ranges that allow users to edit while the sheet is protected
sheet.AddAllowEditRange("EditableRanges", sheet.Range["B4:E12"]);

// Protect the worksheet with a password
sheet.Protect("TestPassword", SheetProtectionType.All);
```

---

# Spire.XLS C# Workbook Protection
## Protect Excel workbook with password
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Protect Workbook with password
workbook.Protect("e-iceblue");
```

---

# Spire.XLS C# Digital Signature Removal
## Remove all digital signatures from an Excel workbook
```csharp
// Remove all digital signatures from the workbook
workbook.RemoveAllDigitalSignatures();
```

---

# spire.xls csharp unprotect worksheet
## unlock password protected Excel sheet
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Unprotect the worksheet with password.
sheet.Unprotect("e-iceblue");
```

---

# spire.xls csharp unlock worksheet
## unlock a simple Excel worksheet
```csharp
// Create a workbook.
Workbook workbook = new Workbook();

// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

// Unlock the worksheet
sheet.Unprotect();
```

---

# spire.xls csharp textbox
## extract text from textbox in excel
```csharp
// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Get the first textbox from the worksheet
XlsTextBoxShape shape = sheet.TextBoxes[0] as XlsTextBoxShape;

// Extract text from the text box
StringBuilder content = new StringBuilder();
content.AppendLine("The text extracted from the TextBox is: ");
content.AppendLine(shape.Text);
```

---

# spire.xls csharp textbox
## get textbox by name in excel worksheet
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the default first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Insert a TextBox at cell A2
sheet.Range["A2"].Text = "Name：";
ITextBoxShape textBox = sheet.TextBoxes.AddTextBox(2, 2, 18, 65);

// Set the name of the TextBox
textBox.Name = "FirstTextBox";

// Set the text for the TextBox
textBox.Text = "Spire.XLS for .NET is a professional Excel .NET component that can be used in any type of .NET 2.0, 3.5, 4.0 or 4.5 framework application, both ASP.NET web sites and Windows Forms application.";

// Get the TextBox by its name
ITextBoxShape FindTextBox = sheet.TextBoxes["FirstTextBox"];

// Get the text content of the TextBox
string text = FindTextBox.Text;
```

---

# spire.xls csharp textbox manipulation
## modify textbox text and alignment in Excel worksheet
```csharp
// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Get the first textbox from the worksheet
ITextBox tb = sheet.TextBoxes[0];

// Change the text of the textbox
tb.Text = "Spire.XLS for .NET";

// Set the alignment of the textbox as center
tb.HAlignment = CommentHAlignType.Center;
tb.VAlignment = CommentVAlignType.Center;
```

---

# Spire.XLS C# Textbox Border Removal
## Remove borderline from textbox in Excel
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Create a new worksheet and add a chart to the worksheet
Worksheet sheet = workbook.Worksheets.Add("Remove Borderline");
Chart chart = sheet.Charts.Add();

// Create textbox1 in the chart and input text information
XlsTextBoxShape textbox1 = chart.TextBoxes.AddTextBox(50, 50, 100, 600) as XlsTextBoxShape;
textbox1.Text = "The solution with borderline";

// Create textbox2 in the chart, input text information, and remove the borderline
XlsTextBoxShape textbox2 = chart.TextBoxes.AddTextBox(1000, 50, 100, 600) as XlsTextBoxShape;
textbox2.Text = "The solution without borderline";
// Remove the border line
textbox2.Line.Weight = 0; 
```

---

# spire.xls csharp textbox
## replace text in excel textbox
```csharp
// Define tags and replacement text
string tag = "TAG_1$TAG_2";
string replace = "Spire.XLS for .NET$Spire.XLS for JAVA";

// Replace text in textboxes
for (int i = 0; i < tag.Split('$').Length; i++)
{
    ReplaceTextInTextBox(sheet, "<" + tag.Split('$')[i] + ">", replace.Split('$')[i]);
}

// Method to replace text in textboxes
private void ReplaceTextInTextBox(Worksheet sheet, string sFind, string sReplace)
{
    for (int i = 0; i < sheet.TextBoxes.Count; i++)
    {
        ITextBox tb = sheet.TextBoxes[i];
        if (!String.IsNullOrEmpty(tb.Text))
        {
            if (tb.Text.Contains(sFind))
            {
                tb.Text = tb.Text.Replace(sFind, sReplace);
            }
        }
    }
}
```

---

# spire.xls csharp textbox formatting
## set font and background for textbox
```csharp
// Get the textbox which will be edited
XlsTextBoxShape shape = sheet.TextBoxes[0] as XlsTextBoxShape;

// Set the font properties for the textbox
ExcelFont font = workbook.CreateFont();
font.FontName = "Century Gothic";
font.Size = 10;
font.IsBold = true;
font.Color = Color.Blue;
(new RichText(shape.RichText)).SetFont(0, shape.Text.Length - 1, font);

// Set the background color for the textbox
shape.Fill.FillType = ShapeFillType.SolidColor;
shape.Fill.ForeKnownColor = ExcelColors.BlueGray;
```

---

# spire.xls csharp textbox
## set internal margin of textbox
```csharp
// Add a textbox to the sheet and set its position and size
XlsTextBoxShape textbox = sheet.TextBoxes.AddTextBox(4, 2, 100, 300) as XlsTextBoxShape;

// Set the text on the textbox
textbox.Text = "Insert TextBox in Excel and set the margin for the text";
textbox.HAlignment = CommentHAlignType.Center;
textbox.VAlignment = CommentVAlignType.Center;

// Set the inner margins of the contents
textbox.InnerLeftMargin = 1;
textbox.InnerRightMargin = 3;
textbox.InnerTopMargin = 1;
textbox.InnerBottomMargin = 1;
```

---

# spire.xls csharp textbox
## Set wrap text for textbox in Excel
```csharp
// Get the text box
XlsTextBoxShape shape = sheet.TextBoxes[0] as XlsTextBoxShape;

// Set wrap text
shape.IsWrapText = true;
```

---

# spire.xls csharp worksheet activation
## activate a specific worksheet in an excel workbook
```csharp
// Get the second worksheet from the workbook
Worksheet sheet = workbook.Worksheets[1];

// Activate the sheet
sheet.Activate();
```

---

# spire.xls csharp page breaks
## add horizontal and vertical page breaks in Excel worksheet
```csharp
// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Add a horizontal page break at cell E4
sheet.HPageBreaks.Add(sheet.Range["E4"]);

// Add a vertical page break at cell C4
sheet.VPageBreaks.Add(sheet.Range["C4"]);
```

---

# spire.xls csharp worksheet
## add new worksheet to workbook
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Add a new worksheet named "AddedSheet"
Worksheet sheet = workbook.Worksheets.Add("AddedSheet");
sheet.Range["C5"].Text = "This is a new sheet.";
```

---

# spire.xls csharp style
## apply style to worksheet
```csharp
// Create a cell style
CellStyle style = workbook.Styles.Add("newStyle");
style.Color = Color.LightBlue;
style.Font.Color = Color.White;
style.Font.Size = 15;
style.Font.IsBold = true;

// Apply the style to the first worksheet
sheet.ApplyStyle(style);
```

---

# spire.xls csharp worksheet copy
## copy multiple worksheets to a single sheet
```csharp
// Get the first worksheet.
Worksheet sheet1 = workbook.Worksheets[0];

// Copy all objects(such as text, shape, image...) from sheet2 to sheet1
for (int i = 1; i < workbook.Worksheets.Count; i++)
{
    Worksheet sheet2 = workbook.Worksheets[i];
    sheet2.Copy((CellRange)sheet2.MaxDisplayRange, sheet1, sheet1.LastRow + 1, sheet2.FirstColumn, true);
}
```

---

# Spire.XLS C# Copy Worksheet
## Copy a worksheet from one Excel file to another
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Create another Workbook
Workbook workbook1 = new Workbook();

// Get the first worksheet in the new workbook
Worksheet sheet1 = workbook1.Worksheets[0];

// Copy the source worksheet to the destination worksheet in the new workbook
sheet1.CopyFrom(sheet);
```

---

# spire.xls csharp worksheet copy
## copy worksheet within workbook
```csharp
// Get the first worksheet and add a new worksheet named "MySheet"
Worksheet sheet = workbook.Worksheets[0];
Worksheet sheet1 = workbook.Worksheets.Add("MySheet");

// Get the source range from the first worksheet
CellRange sourceRange = sheet.AllocatedRange;

// Copy the content of the source range to the second worksheet
sheet.Copy(sourceRange, sheet1, sheet.FirstRow, sheet.FirstColumn, true);
```

---

# Spire.XLS C# Worksheet Operations
## Copy only visible worksheets from one workbook to another
```csharp
// Create a new workbook to copy visible sheets
Workbook workbookNew = new Workbook();
workbookNew.Version = ExcelVersion.Version2013;
workbookNew.Worksheets.Clear();

// Loop through the worksheets in the original workbook
foreach (Worksheet sheet in workbook.Worksheets)
{
    // Check if the worksheet is visible
    if (sheet.Visibility == WorksheetVisibility.Visible)
    {
        // Copy the visible sheet to the new workbook
        workbookNew.Worksheets.AddCopy(sheet);
    }
}
```

---

# spire.xls csharp worksheet copy
## copy worksheet from one workbook to another
```csharp
// Create source and target workbooks
Workbook sourceWorkbook = new Workbook();
Workbook targetWorkbook = new Workbook();

// Get the first worksheet from the source workbook
Worksheet srcWorksheet = sourceWorkbook.Worksheets[0];

// Add a new worksheet to the target workbook
Worksheet targetWorksheet = targetWorkbook.Worksheets.Add("added");

// Copy the source worksheet to the target worksheet
targetWorksheet.CopyFrom(srcWorksheet);
```

---

# Spire.XLS C# Detect Empty Worksheets
## This code demonstrates how to detect if a worksheet in an Excel workbook is empty using the IsEmpty property
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Get the first worksheet from the workbook
Worksheet worksheet1 = workbook.Worksheets[0];

// Detect if the first worksheet is empty
bool detect1 = worksheet1.IsEmpty;

// Get the second worksheet from the workbook
Worksheet worksheet2 = workbook.Worksheets[1];

// Detect if the second worksheet is empty
bool detect2 = worksheet2.IsEmpty;
```

---

# spire.xls csharp worksheet data filling
## Core functionality for filling data into an Excel worksheet using Spire.XLS
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];

// Fill data
worksheet.Range["A1"].Style.Font.IsBold = true;
worksheet.Range["B1"].Style.Font.IsBold = true;
worksheet.Range["C1"].Style.Font.IsBold = true;
worksheet.Range["A1"].Text = "Month";
worksheet.Range["A2"].Text = "January";
worksheet.Range["A3"].Text = "February";
worksheet.Range["A4"].Text = "March";
worksheet.Range["A5"].Text = "April";
worksheet.Range["B1"].Text = "Payments";
worksheet.Range["B2"].NumberValue = 251;
worksheet.Range["B3"].NumberValue = 515;
worksheet.Range["B4"].NumberValue = 454;
worksheet.Range["B5"].NumberValue = 874;
worksheet.Range["C1"].Text = "Sample";
worksheet.Range["C2"].Text = "Sample1";
worksheet.Range["C3"].Text = "Sample2";
worksheet.Range["C4"].Text = "Sample3";
worksheet.Range["C5"].Text = "Sample4";

//Set width for the second column
worksheet.SetColumnWidth(2, 10);
```

---

# Spire.XLS C# Freeze Panes
## Freeze top row in Excel worksheet
```csharp
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Freeze Top Row
sheet.FreezePanes(2,1);
```

---

# Spire.XLS C# Get Custom Properties
## Retrieve custom properties of an Excel worksheet
```csharp
// Get the first sheet
Worksheet worksheet = workbook.Worksheets[0];

// Get the custom properties of the first sheet
ICustomPropertiesCollection customProperties = worksheet.CustomProperties;
for (int i = 0; i < customProperties.Count; i++)
{
    XlsCustomProperty xcp = customProperties[i];
    string name = xcp.Name;
    string value = xcp.Value;
}
```

---

# spire.xls csharp get freeze pane range
## retrieve freeze pane information from an excel worksheet
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
int rowIndex;
int colIndex;

// The row and column index of the frozen pane is passed through the out parameter.
// If it returns to 0, it means that it is not frozen
sheet.GetFreezePanes(out rowIndex, out colIndex);

string range = "Row index: " + rowIndex + ", column index: " + colIndex;
```

---

# Spire.XLS C# Font Extraction
## Extract list of fonts used in an Excel workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();
// Load a excel document
workbook.LoadFromFile("templateAz.xlsx");

List<ExcelFont> fonts = new List<ExcelFont>();

// Loop all sheets of workbook
foreach (Worksheet sheet in workbook.Worksheets)
{
    for (int r = 0; r < sheet.Rows.Length; r++)
    {
        for (int c = 0; c < sheet.Rows[r].CellList.Count; c++)
        {
            //Get the font of cell and add it to list
            fonts.Add(sheet.Rows[r].CellList[c].Style.Font);
        }
    }
}
```

---

# spire.xls csharp page count
## get worksheet page count
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get Split Page Info
var pageInfoList = workbook.GetSplitPageInfo();

// Get page count
for (int i = 0; i < workbook.Worksheets.Count; i++)
{
    string sheetname = workbook.Worksheets[i].Name;
    int pagecount = pageInfoList[i].Count;
    // Process page count for each worksheet
}
```

---

# spire.xls csharp get paper size
## Retrieve page dimensions from Excel worksheets
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Iterate through worksheets
foreach (Worksheet sheet in workbook.Worksheets)
{
    // Get page width
    double width = sheet.PageSetup.PageWidth;

    // Get page height
    double height = sheet.PageSetup.PageHeight;
}
```

---

# spire.xls csharp get worksheet names
## Extract names of all worksheets from an Excel workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the names of all worksheets
StringBuilder sb = new StringBuilder();
foreach(Worksheet sheet in workbook.Worksheets)
{
    sb.AppendLine(sheet.Name);
}
```

---

# spire.xls csharp worksheet visibility
## hide or show excel worksheets
```csharp
// Hide the sheet named "Sheet1"
workbook.Worksheets["Sheet1"].Visibility = WorksheetVisibility.Hidden;

// Show the second sheet
workbook.Worksheets[1].Visibility = WorksheetVisibility.Visible;
```

---

# spire.xls csharp hide worksheet tabs
## hide excel worksheet tabs using ShowTabs property
```csharp
//Hide worksheet tab
workbook.ShowTabs = false;
```

---

# spire.xls csharp hide zero values
## Hide zero values in Excel worksheet
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];

// Set false to hide the zero values in sheet
sheet.IsDisplayZeros = false;
```

---

# spire.xls csharp link content property
## create and link custom document property to content
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Add a custom document property
workbook.CustomDocumentProperties.Add("Test", "MyNamedRange");

// Get the added document property
ICustomDocumentProperties properties = workbook.CustomDocumentProperties;
DocumentProperty property = (DocumentProperty)properties["Test"];

// Link to content 
property.LinkToContent = true;
```

---

# spire.xls csharp move chartsheet
## move chartsheets within excel workbook
```csharp
// Move the first chartsheet to the position of the third sheet (including chartsheet and worksheet) 
workbook.Chartsheets[0].MoveSheet(2);

// Move the first sheet to the position of the first chartsheet
workbook.Chartsheets[0].MoveChartsheet(0);
```

---

# spire.xls csharp worksheet
## move worksheet to a specific position
```csharp
//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Move worksheet
sheet.MoveWorksheet(2);
```

---

# spire.xls csharp page break preview
## set page break preview mode zoom scale
```csharp
//Create a workbook.
Workbook workbook = new Workbook();

//Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];

//Set the scale of PageBreakView mode in Excel file.
sheet.ZoomScalePageBreakView = 80;
```

---

# spire.xls csharp page breaks
## remove page breaks in excel worksheets
```csharp
// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];

// Clear all the vertical page breaks
sheet.VPageBreaks.Clear();

// Remove the first horizontal Page Break
sheet.HPageBreaks.RemoveAt(0);

// Set the ViewMode as Preview to see how the page breaks work
sheet.ViewMode = ViewMode.Preview;
```

---

# spire.xls csharp worksheet
## remove worksheet from workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Remove a worksheet by sheet index
workbook.Worksheets.RemoveAt(1);
```

---

# Spire.XLS C# Page Break
## Set page breaks in Excel worksheet
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Set Excel Page Break Horizontally
sheet.HPageBreaks.Add(sheet.Range["A8"]);
sheet.HPageBreaks.Add(sheet.Range["A14"]);

// Set Excel Page Break Vertically
//sheet.VPageBreaks.Add(sheet.Range["B1"]);
//sheet.VPageBreaks.Add(sheet.Range["C1"]);

// Set view mode to Preview mode
workbook.Worksheets[0].ViewMode = ViewMode.Preview;
```

---

# spire.xls csharp worksheet tab color
## Set tab color for Excel worksheets
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Get the first worksheet
Worksheet worksheet = workbook.Worksheets[0];

//Set the tab color of first sheet to be red 
worksheet.TabColor = Color.Red;

//Set the tab color of second sheet to be green 
worksheet = workbook.Worksheets[1];
worksheet.TabColor = Color.Green;

//Set the tab color of third sheet to be blue 
worksheet = workbook.Worksheets[2];
worksheet.TabColor = Color.LightBlue;
```

---

# spire.xls csharp worksheet view mode
## Set worksheet view mode to Preview
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Set the view mode as Preview
workbook.Worksheets[0].ViewMode = ViewMode.Preview;
```

---

# Spire.XLS C# Grid Lines Control
## Show or hide grid lines in Excel worksheets
```csharp
// Get the first and second worksheet
Worksheet sheet1 = workbook.Worksheets[0];
Worksheet sheet2 = workbook.Worksheets[1];

// Hide grid line in the first worksheet
sheet1.GridLinesVisible = false;

// Show grid line in the second worksheet
sheet2.GridLinesVisible = true;
```

---

# spire.xls csharp show worksheet tabs
## show or hide worksheet tabs in excel workbook
```csharp
// Show worksheet tab
workbook.ShowTabs = true;
```

---

# spire.xls csharp split worksheet
## split worksheet into multiple panes
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Vertical and horizontal split the worksheet into four panes
sheet.FirstVisibleColumn = 2;
sheet.FirstVisibleRow = 5;
sheet.VerticalSplit = 4000;
sheet.HorizontalSplit = 5000;

// Set the active pane
sheet.ActivePane = 1;
```

---

# Spire.XLS C# Unfreeze Panes
## This code demonstrates how to unfreeze panes in an Excel worksheet using Spire.XLS library
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Unfreeze the panes
sheet.RemovePanes();
```

---

# spire.xls worksheet protection verification
## verify if worksheet is password protected
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the Excel document from disk
workbook.LoadFromFile("ProtectedWorksheet.xlsx");

// Get the first worksheet
Worksheet worksheet = workbook.Worksheets[0];

// Verify the first worksheet 
bool detect = worksheet.IsPasswordProtected;

// Set string format for displaying
string result = string.Format("The first worksheet is password protected or not: " + detect);
```

---

# spire.xls csharp zoom factor
## set worksheet zoom factor
```csharp
//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

//Set the zoom factor of the sheet to 85
sheet.Zoom = 85;
```

---

# spire.xls csharp document properties
## access document properties from excel workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the document from disk
workbook.LoadFromFile("sample.xlsx");

// Create string builder
StringBuilder builder = new StringBuilder();

// Get all document properties
ICustomDocumentProperties properties = workbook.CustomDocumentProperties;

// Access document property by property name
DocumentProperty property1 = (DocumentProperty)properties["Editor"];
builder.AppendLine(property1.Name + " " + property1.Value);

// Access document property by property index
DocumentProperty property2 = (DocumentProperty)properties[0];
builder.AppendLine(property2.Name + " " + property2.Value);
```

---

# Spire.XLS C# Custom Properties
## Add custom properties to Excel workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Add custom property to mark as final
workbook.CustomDocumentProperties.Add("_MarkAsFinal", true);

// Add other custom properties to the workbook
workbook.CustomDocumentProperties.Add("The Editor", "E-iceblue");
workbook.CustomDocumentProperties.Add("Phone number", 81705109);
workbook.CustomDocumentProperties.Add("Revision number", 7.12);
workbook.CustomDocumentProperties.Add("Revision date", DateTime.Now);
```

---

# spire.xls csharp decrypt workbook
## decrypt password protected Excel workbook
```csharp
// Detect if the Excel workbook is password protected
bool value = Workbook.IsPasswordProtected(fileName);

if (value)
{
    // Load a file with the password specified
    Workbook workbook = new Workbook();
    workbook.OpenPassword = "eiceblue";
    workbook.LoadFromFile(fileName);

    // Decrypt workbook
    workbook.UnProtect();

    // Save the document
    workbook.SaveToFile("DecryptWorkbook_result.xlsx", ExcelVersion.Version2010);

    // Dispose of the workbook object to release resources
    workbook.Dispose();
}
```

---

# spire.xls csharp detect excel version
## Detect Excel file version using Spire.XLS
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the document
workbook.LoadFromFile(file);

// Get the version
ExcelVersion version = workbook.Version;

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls csharp vba detection
## detect VBA macros in Excel workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the document from disk
workbook.LoadFromFile("MacroSample.xls");

// Detect if the Excel file contains VBA macros
bool hasMacros = workbook.HasMacros;
```

---

# Spire.XLS C# Disable DTD
## Disables DTD processing in Excel workbooks
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Disable DTD
workbook.ProhibitDtd = true;
```

---

# spire.xls csharp encrypt workbook
## protect workbook with password
```csharp
// Create a workbook 
Workbook workbook = new Workbook();

// Protect Workbook with the password you want
workbook.Protect("eiceblue");
```

---

# spire.xls csharp get document properties
## retrieve built-in and custom properties from excel workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the document from disk
workbook.LoadFromFile("WorksheetSample1.xlsx");

// Get the general excel properties
BuiltInDocumentProperties properties1 = workbook.DocumentProperties;
StringBuilder sb = new StringBuilder();
sb.AppendLine("Excel Properties:");
for (int i = 0; i < properties1.Count; i++)
{
    // Get property name
    string name = properties1[i].Name;
    // Get property vaule
    string value = properties1[i].Value.ToString();
    sb.AppendLine(name + ": " + value);
}
sb.AppendLine();

//Get the custom properties
ICustomDocumentProperties properties2 = workbook.CustomDocumentProperties;
sb.AppendLine("Custom Properties:");
for (int i = 0; i < properties2.Count; i++)
{
    // Get property name
    string name = properties2[i].Name;
    // Get property vaule
    string value = properties2[i].Value.ToString();
    sb.AppendLine(name + ": " + value);
}

//Save the document
string output = "GetProperties.txt";
File.WriteAllText(output, sb.ToString());

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS C# Hide Window
## Hide Excel window using Spire.XLS library
```csharp
// Create a workbook
Workbook workbook = new Workbook();
// Hide window
workbook.IsHideWindow = true;
```

---

# spire.xls c# load save excel with macros
## demonstrates how to load an excel file with macros, modify it, and save it back while preserving the macros
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MacroSample.xls");

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Set value for cell A5
sheet.Range["A5"].Text = "This is a simple test!";

// Save the document
string output = "LoadAndSaveFileWithMacro.xls";
workbook.SaveToFile(output, ExcelVersion.Version97to2003);
```

---

# spire.xls csharp merge excel files
## merge multiple excel files into one workbook
```csharp
// Create a new workbook
Workbook newbook = new Workbook();
newbook.Version = ExcelVersion.Version2013;

// Clear all worksheets
newbook.Worksheets.Clear();

// Create a temporary workbook
Workbook tempbook = new Workbook();

foreach (string file in files)
{
    // Load the file
    tempbook.LoadFromFile(file);
    foreach (Worksheet sheet in tempbook.Worksheets)
    {
        // Copy every sheet in a workbook
        newbook.Worksheets.AddCopy(sheet, WorksheetCopyType.CopyAll);
    }
    // Dispose of the workbook object to release resources
    tempbook.Dispose();
}

// Save the file
newbook.SaveToFile("MergeExcelFiles.xlsx", ExcelVersion.Version2010);

// Dispose of the workbook object to release resources 
newbook.Dispose();
```

---

# spire.xls csharp encrypted file
## open encrypted Excel file by trying multiple passwords
```csharp
// Password array 
String[] passwords = new String[4] { "password1", "password2", "password3", "1234" };

for (int i = 0; i < passwords.Length; i++)
{
    try
    {
        // Create a workbook
        Workbook workbook = new Workbook();

        // Set open password
        workbook.OpenPassword = passwords[i];

        //Load the encrypted document
        workbook.LoadFromFile(/* encrypted file path */);
        
        // Dispose of the workbook object to release resources 
        workbook.Dispose();
    }
    catch (Exception ex)
    {
        // Password is not correct
    }
}
```

---

# spire.xls csharp open different file formats
## demonstrates how to open various Excel file formats using Spire.XLS
```csharp
// 1. Load file by file path
// Create a workbook
Workbook workbook1 = new Workbook();
// Load the document from disk
workbook1.LoadFromFile(filepath);

// Dispose of the workbook object to release resources 
workbook1.Dispose();

// 2. Load file by file stream
FileStream stream = new FileStream(filepath, FileMode.Open);

// Create a workbook
Workbook workbook2 = new Workbook();

// Load the document from disk
workbook2.LoadFromStream(stream);
stream.Dispose();

// Dispose of the workbook object to release resources 
workbook2.Dispose();

// 3. Open Microsoft Excel 97 - 2003 file
Workbook wbExcel97 = new Workbook();
wbExcel97.LoadFromFile(filepath97, ExcelVersion.Version97to2003);

// 4. Open xml file
Workbook wbXML = new Workbook();
wbXML.LoadFromXml(filepathXml);

// Dispose of the workbook object to release resources 
wbExcel97.Dispose();

//5. Open csv file
Workbook wbCSV = new Workbook();
wbCSV.LoadFromFile(filepathCsv, ",", 1, 1);

// Dispose of the workbook object to release resources 
wbCSV.Dispose();
```

---

# spire.xls csharp stream
## read excel workbook from stream
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Open excel from a stream
FileStream fileStream = File.OpenRead(@"..\..\..\..\..\..\Data\ReadStream.xlsx");
fileStream.Seek(0, SeekOrigin.Begin);

// Load file from stream
workbook.LoadFromStream(fileStream);
```

---

# Spire.XLS C# Remove Custom Properties
## Remove custom document properties from an Excel workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Retrieve a list of all custom document properties of the Excel file
ICustomDocumentProperties customDocumentProperties = workbook.CustomDocumentProperties;

// Remove "Editor" custom document property
customDocumentProperties.Remove("Editor");

// Dispose of the workbook object to release resources 
workbook.Dispose();
```

---

# spire.xls csharp save files
## save workbook in different file formats
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

// Save in Excel 97-2003 format
workbook.SaveToFile("result.xls",ExcelVersion.Version97to2003);

// Save in Excel2010 xlsx format
workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010);

// Save in XLSB format
workbook.SaveToFile("result.xlsb", ExcelVersion.Xlsb2010);

// Save in ODS format
workbook.SaveToFile("result.ods", ExcelVersion.ODS);

// Save in PDF format
workbook.SaveToFile("result.pdf", FileFormat.PDF);

// Save in XML format
workbook.SaveToFile("result.xml",FileFormat.XML);

// Save in XPS format
workbook.SaveToFile("result.xps", FileFormat.XPS);

// Dispose of the workbook object to release resources 
workbook.Dispose();
```

---

# Spire.XLS C# Save to Stream
## Save Excel workbook to file stream
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Load the document from disk
workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SaveStream.xls");

// Save an excel workbook to stream
FileStream fileStream = new FileStream("SaveStream.xlsx", FileMode.Create);
workbook.SaveToStream(fileStream, FileFormat.Version2010);

// Close the stream
fileStream.Close();

// Dispose of the workbook object to release resources 
workbook.Dispose();
```

---

# spire.xls csharp excel calculation mode
## Set Excel calculation mode to manual
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Set excel calculation mode as Manual
workbook.CalculationMode = ExcelCalculationMode.Manual;
```

---

# spire.xls csharp margins
## Set page margins for Excel worksheet
```csharp
// Set margins for top, bottom, left and right, here the unit of measure is Inch
sheet.PageSetup.TopMargin = 0.3;
sheet.PageSetup.BottomMargin = 1;
sheet.PageSetup.LeftMargin = 0.2;
sheet.PageSetup.RightMargin = 1;

// Set the header margin and footer margin
sheet.PageSetup.HeaderMarginInch = 0.1;
sheet.PageSetup.FooterMarginInch = 0.5;
```

---

# Spire.XLS C# Theme Setting
## Set theme color in Excel workbook
```csharp
// Create a workbook
Workbook srcWorkbook = new Workbook();

// Get the first worksheet
Worksheet srcWorksheet = srcWorkbook.Worksheets[0];

// Create an empty workbook
Workbook workbook = new Workbook();
workbook.Worksheets.Clear();
workbook.Worksheets.AddCopy(srcWorksheet);

// 1. Copy the theme of the workbook
//workbook.CopyTheme(srcWorkbook);

// 2. Set a certain type of color of the default theme in the workbook
workbook.SetThemeColor(ThemeColorType.Dk1, Color.SkyBlue);
```

---

# spire.xls csharp track changes
## accept or reject tracked changes in excel workbook
```csharp
// Create a workbook
Workbook workbook = new Workbook();

// Accept the changes or reject the changes.
//workbook.AcceptAllTrackedChanges();
workbook.RejectAllTrackedChanges();
```

---

# Spire.XLS C# Track Changes
## Enable track changes in Excel workbook
```csharp
// Create a new workbook object
Workbook workbook = new Workbook();

// Enable track changes 
workbook.TrackedChanges = true;
```

---

# spire.xls csharp data export
## export data while preserving data types
```csharp
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Export DataTable with data type preservation
ExportTableOptions options = new ExportTableOptions();
options.ExportColumnNames = true;
options.KeepDataFormat = false;
options.KeepDataType = true;
options.RenameStrategy = RenameStrategy.Digit;

// Export data to data table
DataTable table = sheet.ExportDataTable(1, 1, sheet.LastDataRow, sheet.LastDataColumn, options);
```

---

# Remove Duplicated Rows in Excel
## Remove duplicate rows from an Excel worksheet using Spire.XLS
```csharp
// Remove duplicated rows in the worksheet
sheet.RemoveDuplicates();

// Remove the duplicate rows within the specified range
// sheet.RemoveDuplicates(int startRow, int startColumn, int endRow, int endColumn);
// Remove the duplicated rows based on specific columns and headers
// sheet.RemoveDuplicates(int startRow, int startColumn, int endRow, int endColumn, boolean hasHeaders, int[] columnOffsets)
```

---

# Markdown to XLSX Conversion
## Convert Markdown files to Excel XLSX format using Spire.Xls library
```csharp
// Create a new Workbook instance
Workbook workbook = new Workbook();

// Load content from a Markdown file into the workbook
workbook.LoadFromMarkdown(markdownFilePath);

// Save the workbook to an Excel file
workbook.SaveToFile(outputFileName, ExcelVersion.Version2016);

// Release the resources used by the workbook object
workbook.Dispose();
```

---

# Excel Shape Hyperlink
## Add hyperlink to shapes in Excel worksheet
```csharp
// Get the reference to the first sheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Get all the shapes in the sheet
PrstGeomShapeCollection prstGeomShapeType = sheet.PrstGeomShapes;

// Set the hyperlink for each shape
for (int i = 0; i < prstGeomShapeType.Count; i++)
{
    // Get the shape
    XlsPrstGeomShape shape = (XlsPrstGeomShape)prstGeomShapeType[i];

    // Set the hyperlink address
    shape.HyLink.Address = "https://www.e-iceblue.com/Download/download-excel-for-net-now.html";
}
```

---

# spire.xls csharp pivot table
## create pivot table group by value
```csharp
// Get the reference to the first sheet in the workbook
Worksheet pivotSheet = workbook.Worksheets[0];

// Cast the first PivotTable in the PivotTables collection to an XlsPivotTable object.
XlsPivotTable pivot = (XlsPivotTable)pivotSheet.PivotTables[0];

// Retrieve the PivotField named "number" from the PivotTable and cast it to a PivotField object.
PivotField dateBaseField = pivot.PivotFields["number"] as PivotField;

// Create a group for the PivotField, starting at 3000, ending at 3800, with an interval of 1.
dateBaseField.CreateGroup(3000, 3800, 1);

// Recalculate the data in the PivotTable to reflect the changes made.
pivot.CalculateData();
```

---

# spire.xls csharp pivot table slicer
## create and configure slicers from pivot table
```csharp
// Get pivot table collection
Spire.Xls.Collections.PivotTablesCollection pivotTables = worksheet.PivotTables;

//Add a PivotTable to the worksheet
CellRange dataRange = worksheet.Range["A1:C9"];
PivotCache cache = wb.PivotCaches.Add(dataRange);

//Cell to put the pivot table
Spire.Xls.PivotTable pt = worksheet.PivotTables.Add("TestPivotTable", worksheet.Range["A12"], cache);

//Drag the fields to the row area.
PivotField pf = pt.PivotFields["fruit"] as PivotField;
pf.Axis = AxisTypes.Row;
PivotField pf2 = pt.PivotFields["year"] as PivotField;
pf2.Axis = AxisTypes.Column;

//Drag the field to the data area.
pt.DataFields.Add(pt.PivotFields["amount"], "SUM of Count", SubtotalTypes.Sum);

//Set PivotTable style
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium10;

pt.CalculateData();

//Get slicer collection
XlsSlicerCollection slicers = worksheet.Slicers;

int index = slicers.Add(pt, "E12", 0);

XlsSlicer xlsSlicer = slicers[index];
xlsSlicer.Name = "xlsSlicer";
xlsSlicer.Width = 100;
xlsSlicer.Height = 120;
xlsSlicer.StyleType = SlicerStyleType.SlicerStyleLight2;
xlsSlicer.PositionLocked = true;

//Get SlicerCache object of current slicer
XlsSlicerCache slicerCache = xlsSlicer.SlicerCache;
slicerCache.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithNoData;

//Style setting
XlsSlicerCacheItemCollection slicerCacheItems = xlsSlicer.SlicerCache.SlicerCacheItems;
XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems[0];
xlsSlicerCacheItem.Selected = false;

XlsSlicerCollection slicers_2 = worksheet.Slicers;

IPivotField r1 = pt.PivotFields["year"];
int index_2 = slicers_2.Add(pt, "I12", r1);

XlsSlicer xlsSlicer_2 = slicers[index_2];
xlsSlicer_2.RowHeight = 40;
xlsSlicer_2.StyleType = SlicerStyleType.SlicerStyleLight3;
xlsSlicer_2.PositionLocked = false;

//Get SlicerCache object of current slicer
XlsSlicerCache slicerCache_2 = xlsSlicer_2.SlicerCache;
slicerCache_2.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithDataAtTop;

//Style setting
XlsSlicerCacheItemCollection slicerCacheItems_2 = xlsSlicer_2.SlicerCache.SlicerCacheItems;
XlsSlicerCacheItem xlsSlicerCacheItem_2 = slicerCacheItems_2[1];
xlsSlicerCacheItem_2.Selected = false;
pt.CalculateData();
```

---

# spire.xls csharp slicer
## create slicers from table in excel
```csharp
// Get slicer collection
XlsSlicerCollection slicers = worksheet.Slicers;

//Create a table with the data from the specific cell range.
IListObject table = worksheet.ListObjects.Create("Super Table", worksheet.Range["A1:C9"]);

int count = 3;
int index = 0;
foreach (SlicerStyleType type in Enum.GetValues(typeof(SlicerStyleType)))
{
    count += 5;
    String range = "E" + count;
    index = slicers.Add(table, range.ToString(), 0);

    //Style setting
    XlsSlicer xlsSlicer = slicers[index];
    xlsSlicer.Name = "slicers_" + count;
    xlsSlicer.StyleType = type;
}
```

---

# spire.xls csharp slicer modification
## modify Excel slicer properties including style, caption, and filtering behavior
```csharp
// Get the first worksheet in the workbook
Worksheet worksheet = wb.Worksheets[0];

// Get the slicer collection from the worksheet
XlsSlicerCollection slicers = worksheet.Slicers;

// Get the first slicer from the slicer collection
XlsSlicer xlsSlicer = slicers[0];

// Set the style of the slicer to a dark theme (style type 4)
xlsSlicer.StyleType = SlicerStyleType.SlicerStyleDark4;

// Change the caption (title) of the slicer
xlsSlicer.Caption = "Modified Slicer";

// Lock the position of the slicer to prevent it from being moved in the worksheet
xlsSlicer.PositionLocked = true;

// Get the collection of cache items associated with the slicer
XlsSlicerCacheItemCollection slicerCacheItems = xlsSlicer.SlicerCache.SlicerCacheItems;

// Get the first cache item in the collection
XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems[0];

// Deselect the cache item
xlsSlicerCacheItem.Selected = false;

// Get the display value of the cache item
string displayValue = xlsSlicerCacheItem.DisplayValue;

// Get the slicer cache associated with the slicer
XlsSlicerCache slicerCache = xlsSlicer.SlicerCache;

// Set the cross-filter type to show items even if they have no associated data
slicerCache.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithNoData;
```

---

# Spire.XLS C# Slicer Information Reader
## Read and extract information about slicers in an Excel file
```csharp
// Load Excel file and get worksheet
Workbook wb = new Workbook();
wb.LoadFromFile("SlicerTemplate.xlsx");
Worksheet worksheet = wb.Worksheets[0];

// Get slicer collection
XlsSlicerCollection slicers = worksheet.Slicers;

// Iterate through each slicer
for (int i = 0; i < slicers.Count; i++)
{
    XlsSlicer xlsSlicer = slicers[i];
    
    // Get slicer properties
    string slicerName = xlsSlicer.Name;
    string slicerCaption = xlsSlicer.Caption;
    int numberOfColumns = xlsSlicer.NumberOfColumns;
    double columnWidth = xlsSlicer.ColumnWidth;
    double rowHeight = xlsSlicer.RowHeight;
    bool showCaption = xlsSlicer.ShowCaption;
    bool positionLocked = xlsSlicer.PositionLocked;
    double width = xlsSlicer.Width;
    double height = xlsSlicer.Height;

    // Get slicer cache
    XlsSlicerCache slicerCache = xlsSlicer.SlicerCache;
    
    // Get slicer cache properties
    string sourceName = slicerCache.SourceName;
    bool isTabular = slicerCache.IsTabular;
    string cacheName = slicerCache.Name;

    // Get slicer cache items
    XlsSlicerCacheItemCollection slicerCacheItems = slicerCache.SlicerCacheItems;
    XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems[1];
    
    // Get slicer cache item properties
    bool isSelected = xlsSlicerCacheItem.Selected;
}

// Clean up
wb.Dispose();
```

---

# spire.xls remove slicer
## remove slicers from excel worksheet
```csharp
// Get the slicer collection from the worksheet
XlsSlicerCollection slicers = worksheet.Slicers;

// Example: Remove the first slicer in the collection 
// slicers.RemoveAt(0);

// Clear all slicers from the collection
slicers.Clear();
```

---



