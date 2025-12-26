# Spire.XLS C++ Configuration
## Precompiled header setup for Spire.XLS library
```cpp
#pragma once

#pragma comment(lib,"../lib/Spire.Xls.Cpp.lib")

#define DATAPATH L"Data\\"
#define OUTPUTPATH L"Output\\"
```

---

# spire.xls cpp workbook
## create an excel workbook with five sheets
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();
workbook->CreateEmptySheets(5);
for (int i = 0; i < 5; i++)
{
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
    sheet->SetName((L"Sheet" + std::to_wstring(i)).c_str());
}
```

---

# spire.xls cpp create excel
## create an excel file with one sheet and fill data
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();
workbook->CreateEmptySheets(1);
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
for (int row = 1; row <= 10000; row++)
{
	for (int col = 1; col <= 30; col++)
	{
		dynamic_pointer_cast<CellRange>(sheet->GetRange(row, col))->SetText((L"row" + std::to_wstring(row) + L" col" + std::to_wstring(col)).c_str());
	}
}
workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
```

---

# spire.xls cpp batch creation
## create fifty excel files with multiple worksheets
```cpp
for (int n = 0; n < 50; n++)
{
    intrusive_ptr<Workbook> workbook = new Workbook();
    workbook->CreateEmptySheets(5);
    for (int i = 0; i < 5; i++)
    {
        intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
        sheet->SetName((L"Sheet" + std::to_wstring(i)).c_str());
        for (int row = 1; row <= 150; row++)
        {
            for (int col = 1; col <= 50; col++)
            {
                dynamic_pointer_cast<CellRange>(sheet->GetRange(row, col))->SetText((L"row" + std::to_wstring(row) + L" col" + std::to_wstring(col)).c_str());
            }
        }
    }
    workbook->SaveToFile((output_path + L"Workbook" + std::to_wstring(n) + L".xlsx").c_str(), ExcelVersion::Version2010);
    workbook->Dispose();
}
```

---

# spire.xls cpp helloworld
## create a simple Excel file with Hello World text
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();
//Get the first sheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Hello World");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->AutoFitColumns();
```

---

# spire.xls cpp open existing file
## open an existing Excel file and perform basic operations
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();
workbook->LoadFromFile(inputFile.c_str());

//Add a new sheet, named MySheet
intrusive_ptr<Worksheet> sheet = workbook->GetWorksheets()->Add(L"MySheet");

//Get the reference of L"A1" cell from the cells collection of a worksheet
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Hello World");

workbook->Dispose();
```

---

# spire.xls cpp label control
## add label control to excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add a label control
intrusive_ptr<ILabelShape> label = sheet->GetLabelShapes()->AddLabel(10, 2, 30, 200);
label->SetText(L"This is a Label Control");
```

---

# spire.xls cpp listbox
## add listbox control to excel worksheet
```cpp
//Add listbox control
intrusive_ptr<IListBox> listBox = sheet->GetListBoxes()->AddListBox(13, 4, 120, 100);
listBox->SetSelectionType(SelectionType::Single);
listBox->SetSelectedIndex(2);
listBox->SetDisplay3DShading(true);
listBox->SetListFillRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7:A12")));
```

---

# spire xls cpp scrollbar control
## add scroll bar control to excel worksheet
```cpp
//Set a value for range B10
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10"))->SetValue(std::to_wstring(1).c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10"))->GetStyle()->GetFont()->SetIsBold(true);

//Add scroll bar control
intrusive_ptr<IScrollBarShape> scrollBar = sheet->GetScrollBarShapes()->AddScrollBar(10, 3, 150, 20);
scrollBar->SetLinkedCell(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10")));
scrollBar->SetMin(1);
scrollBar->SetMax(150);
scrollBar->SetIncrementalChange(1);
scrollBar->SetDisplay3DShading(true);
```

---

# spire.xls cpp table
## add table with filter to excel worksheet
```cpp
//Create a List Object named in Table.
sheet->GetListObjects()->Create(L"Table", sheet->GetRange(1, 1, sheet->GetLastRow(), sheet->GetLastColumn()));

//Set the BuiltInTableStyle for List object.
intrusive_ptr<IListObjects> ie = sheet->GetListObjects();
int i = ie->GetCount();
ie->GetItem(0)->SetBuiltInTableStyle(TableBuiltInStyles::TableStyleLight9);
```

---

# spire.xls cpp table
## add total row to table
```cpp
//Create a table with the data from the specific cell range.
intrusive_ptr<IListObject> table = sheet->GetListObjects()->Create(L"Table", dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D4")));

//Display total row.
table->SetDisplayTotalRow(true);

//Add a total row.
intrusive_ptr<Spire::Xls::IList<IListObjectColumn>> list = table->GetColumns();
list->GetItem(0)->SetTotalsRowLabel(L"Total");
list->GetItem(1)->SetTotalsCalculation(ExcelTotalsCalculation::Sum);
list->GetItem(2)->SetTotalsCalculation(ExcelTotalsCalculation::Sum);
list->GetItem(3)->SetTotalsCalculation(ExcelTotalsCalculation::Sum);
```

---

# spire.xls cpp formatting
## apply subscript and superscript to Excel cells
```cpp
//Set the rtf value of L"B3" to L"R100-0.06".
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"));
range->GetRichText()->SetText(L"R100-0.06");

//Create a font. Set the IsSubscript property of the font to L"true".
intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
font->SetIsSubscript(true);
font->SetColor(Spire::Xls::Color::GetGreen());

//Set font for specified range of the text in L"B3".
range->GetRichText()->SetFont(4, 8, font);

//Set the rtf value of L"D3" to L"a2 + b2 = c2".
range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D3"));
range->GetRichText()->SetText(L"a2 + b2 = c2");

//Create a font. Set the IsSuperscript property of the font to L"true".
font = workbook->CreateExcelFont();
font->SetIsSuperscript(true);

//Set font for specified range of the text in L"D3".
range->GetRichText()->SetFont(1, 1, font);
range->GetRichText()->SetFont(6, 6, font);
range->GetRichText()->SetFont(11, 11, font);
```

---

# spire.xls cpp font style
## clone Excel font style
```cpp
//Set A1 cell range's CellStyle.
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"style");
style->GetFont()->SetFontName(L"Calibri");
style->GetFont()->SetColor(Spire::Xls::Color::GetRed());
style->GetFont()->SetSize(12);
style->GetFont()->SetIsBold(true);
style->GetFont()->SetIsItalic(true);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetCellStyleName(style->GetName());

//Clone the same style for B2 cell GetRange.
intrusive_ptr<CellStyle> csOrieign = style->clone();
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetText(L"Text2");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetCellStyleName(csOrieign->GetName());

//Clone the same style for C3 cell GetRange and then reset the font color for the text.
intrusive_ptr<CellStyle> csGreen = style->clone();
csGreen->GetFont()->SetColor(Spire::Xls::Color::GetGreen());
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetText(L"Text3");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetCellStyleName(csGreen->GetName());
```

---

# spire.xls cpp range copy
## copy cell range from source to destination
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Specify a destination range 
intrusive_ptr<CellRange> cells = dynamic_pointer_cast<CellRange>(sheet1->GetRange(L"G1:H19"));

//Copy the selected range to destination range 
dynamic_pointer_cast<CellRange>(sheet1->GetRange(L"B1:C19"))->Copy(cells);
```

---

# spire.xls cpp copy data with style
## copy cell range with style attributes
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get a source range (A1:D3).
intrusive_ptr<CellRange> srcRange = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D3"));

//Create a style object.
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"style");

//Specify the font attribute.
style->GetFont()->SetFontName(L"Calibri");

//Specify the shading color.
style->GetFont()->SetColor(Color::GetRed());

//Specify the border attributes.
style->GetBorders()->Get(BordersLineType::EdgeTop)->SetLineStyle(LineStyleType::Thin);
style->GetBorders()->Get(BordersLineType::EdgeTop)->SetColor(Color::GetBlue());
style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Thin);
style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Color::GetBlue());
style->GetBorders()->Get(BordersLineType::EdgeLeft)->SetLineStyle(LineStyleType::Thin);
style->GetBorders()->Get(BordersLineType::EdgeLeft)->SetColor(Color::GetBlue());
style->GetBorders()->Get(BordersLineType::EdgeRight)->SetLineStyle(LineStyleType::Thin);
style->GetBorders()->Get(BordersLineType::EdgeRight)->SetColor(Color::GetBlue());
srcRange->SetCellStyleName(style->GetName());

//Set the destination range
intrusive_ptr<CellRange> destRange = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A12:D14"));

//Copy the range Demo with style
srcRange->Copy(destRange, true, true);
```

---

# spire.xls cpp copy formula values
## copy only formula values from one range to another in excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the copy option
CopyRangeOptions copyOptions = CopyRangeOptions::OnlyCopyFormulaValue;

//Copy range
sheet->Copy(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:C2")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:C5")), copyOptions);
```

---

# spire.xls cpp nested grouping
## create nested row groups in Excel worksheet
```cpp
//Set the summary rows appear above detail rows.
sheet->GetPageSetup()->SetIsSummaryRowBelow(false);

//Group the rows that you want to group.
sheet->GroupByRows(2, 9, false);
sheet->GroupByRows(4, 5, false);
sheet->GroupByRows(8, 9, false);
```

---

# spire.xls cpp data sorting
## sort data in excel worksheet
```cpp
intrusive_ptr<Workbook> workbook = new Workbook();
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

workbook->GetDataSorter()->GetSortColumns()->Add(2, OrderBy::Ascending);
workbook->GetDataSorter()->GetSortColumns()->Add(3, OrderBy::Ascending);
workbook->GetDataSorter()->Sort(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E19")));
```

---

# c++ excel data validation
## implement different types of data validation in excel cells
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Decimal DataValidation
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B11"))->SetText(L"Input Number(3-6):");
intrusive_ptr<CellRange> rangeNumber = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B12"));
//Set the operator for the data validation.
rangeNumber->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);
//Set the value or expression associated with the data validation.
rangeNumber->GetDataValidation()->SetFormula1(L"3");
//The value or expression associated with the second part of the data validation.
rangeNumber->GetDataValidation()->SetFormula2(L"6");
//Set the data validation type.
rangeNumber->GetDataValidation()->SetAllowType(CellDataType::Decimal);
//Set the data validation error message.
rangeNumber->GetDataValidation()->SetErrorMessage(L"Please input correct number!");
//Enable the error.
rangeNumber->GetDataValidation()->SetShowError(true);
rangeNumber->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

//Date DataValidation
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B14"))->SetText(L"Input Date:");
intrusive_ptr<CellRange> rangeDate = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B15"));
rangeDate->GetDataValidation()->SetAllowType(CellDataType::Date);
rangeDate->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);
rangeDate->GetDataValidation()->SetFormula1(L"1/1/1970");
rangeDate->GetDataValidation()->SetFormula2(L"12/31/1970");
rangeDate->GetDataValidation()->SetErrorMessage(L"Please input correct date!");
rangeDate->GetDataValidation()->SetShowError(true);
rangeDate->GetDataValidation()->SetAlertStyle(AlertStyleType::Warning);
rangeDate->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

//TextLength DataValidation
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B17"))->SetText(L"Input Text:");
intrusive_ptr<CellRange> rangeTextLength = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B18"));
rangeTextLength->GetDataValidation()->SetAllowType(CellDataType::TextLength);
rangeTextLength->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::LessOrEqual);
rangeTextLength->GetDataValidation()->SetFormula1(L"5");
rangeTextLength->GetDataValidation()->SetErrorMessage(L"Enter a Valid String!");
rangeTextLength->GetDataValidation()->SetShowError(true);
rangeTextLength->GetDataValidation()->SetAlertStyle(AlertStyleType::Stop);
rangeTextLength->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

sheet->AutoFitColumn(2);
```

---

# spire.xls cpp group management
## expand and collapse grouped rows in Excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Expand the grouped rows with ExpandCollapseFlags set to expand parent
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A16:G19"))->ExpandGroup(GroupByType::ByRows, ExpandCollapseFlags::ExpandParent);

//Collapse the grouped rows
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A10:G12"))->CollapseGroup(GroupByType::ByRows);
```

---

# spire.xls cpp find and replace
## find and replace data in excel cells
```cpp
//Find the "Area" string
auto ranges = sheet->FindAllString(L"Area", false, false);

//Traverse the found ranges
for (int i = 0; i < ranges->GetCount(); i++)
{
    intrusive_ptr<CellRange> cr = ranges->GetItem(i);
    //Replace it with "Area Code"
    cr->SetText(L"Area Code");
    //Highlight the color
    cr->GetStyle()->SetColor(Spire::Xls::Color::GetYellow());
}
```

---

# spire.xls cpp find data
## find string and number in specific range
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Specify a range
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(1, 1, 12, 8));

//Find string from this range
auto textRanges = range->FindAllString(L"E-iceblue", false, false);

//Check if any text was found
if (textRanges->GetCount() != 0)
{
    for (int i = 0; i < textRanges->GetCount(); i++)
    {
        intrusive_ptr<CellRange> cr = textRanges->GetItem(i);
        wstring address = cr->GetRangeAddress();
    }
}

//Find number from this range
auto ranges = range->FindAllNumber(100, true);

//Check if any number was found
if (ranges->GetCount() != 0)
{
    for (int i = 0; i < ranges->GetCount(); i++)
    {
        intrusive_ptr<CellRange> r = ranges->GetItem(i);
        wstring address = r->GetRangeAddress();
    }
}
```

---

# spire.xls cpp find
## find string and number in Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Find cells with the input string
auto ranges = sheet->FindAllString(L"E-iceblue", false, false);

//Create a string builder
std::wstring* builder = new std::wstring();

//Append the address of found cells in builder
if (ranges->GetCount() != 0)
{
    for (int i = 0; i < ranges->GetCount(); i++)
    {
        intrusive_ptr<CellRange> cr = ranges->GetItem(i);
        wstring address = cr->GetRangeAddress();
        builder->append(L"The address of found text cell is: " + address);
    }
}
else
{
    builder->append(L"No cells that contain the text");
}

//Find cells with the input integer or double
auto numberRanges = sheet->FindAllNumber(100, true);

//Append the address of found cells in builder
if (numberRanges->GetCount() != 0)
{
    for (int i = 0; i < numberRanges->GetCount(); i++)
    {
        intrusive_ptr<CellRange> cr = numberRanges->GetItem(i);
        wstring address = cr->GetRangeAddress();
        builder->append(L"The address of found number cell is: " + address);
    }
}
else
{
    builder->append(L"No cells that contain the number");
}
```

---

# spire.xls cpp import data
## import data from array list to worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Create an empty worksheet
workbook->CreateEmptySheets(1);

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create an ArrayList object
vector<LPCWSTR_S> list;

//Add strings in list
list.push_back(L"Spire.Doc for C++");
list.push_back(L"Spire.XLS for C++");
list.push_back(L"Spire.PDF for C++");
list.push_back(L"Spire.Presentation for C++");

//Insert arrary list in worksheet 
sheet->InsertArray(list, 1, 1, true);
```

---

# spire xls cpp insert controls
## insert form controls into excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add a textbox 
intrusive_ptr<ITextBoxShape> textbox = sheet->GetTextBoxes()->AddTextBox(9, 2, 25, 100);
textbox->SetText(L"Hello World");

//Add a checkbox 
intrusive_ptr<ICheckBox> cb = sheet->GetCheckBoxes()->AddCheckBox(11, 2, 15, 100);
cb->SetCheckState(Spire::Xls::CheckState::Checked);
cb->SetText(L"Check Box 1");

//Add a RadioButton 
intrusive_ptr<IRadioButton> rb = sheet->GetRadioButtons()->Add(13, 2, 15, 100);
rb->SetText(L"Option 1");

//Add a combox
intrusive_ptr<IComboBoxShape> cbx = dynamic_pointer_cast<IComboBoxShape>(sheet->GetComboBoxes()->AddComboBox(15, 2, 15, 100));
cbx->SetListFillRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A41:A47")));
```

---

# spire.xls cpp html string
## insert HTML string into Excel cell
```cpp
// Create a Workbook object
intrusive_ptr<Workbook> workbook = new Workbook();

// Get the first worksheet from the workbook
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Define an HTML code to be placed in cell A1
std::wstring htmlCode = L"<div>first line<br>second line<br>third line</div>";

// Get the cell range for cell A1 and set the HTML string
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"));
range->SetHtmlString(htmlCode.c_str());
```

---

# spire.xls cpp named ranges
## create and configure named range in Excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Creating a named range
intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Add(L"NewNamedRange");
//Setting the range of the named range
NamedRange->SetRefersToRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:E12")));
```

---

# spire.xls cpp replace and highlight
## replace text and highlight cells in Excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

auto ranges = sheet->FindAllString(L"Total", true, true);
for (int i = 0; i < ranges->GetCount(); i++)
{
    intrusive_ptr<CellRange> cr = ranges->GetItem(i);
    //reset the text, in other words, replace the text
    cr->SetText(L"Sum");
    //set the color
    cr->GetStyle()->SetColor(Spire::Xls::Color::GetYellow());
}
```

---

# spire.xls cpp data extraction
## retrieve and extract specific rows from excel worksheet
```cpp
// Create a new workbook instance and get the first worksheet.
intrusive_ptr<Workbook> newBook = new Workbook();
intrusive_ptr<Worksheet> newSheet = dynamic_pointer_cast<Worksheet>(newBook->GetWorksheets()->Get(0));

//Create a new workbook instance and load the sample Excel file.
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet.
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Retrieve data and extract it to the first worksheet of the new excel workbook.
int k = 1;
int columnCount = sheet->GetColumns()->GetCount();
for (int i = 0; i < sheet->GetColumns()->GetItem(0)->GetCells()->GetCount(); i++)
{
	intrusive_ptr<XlsRange> range = sheet->GetColumns()->GetItem(0)->GetCells()->GetItem(i);
	if (wcscmp(range->GetText(), L"teacher") == 0)
	{
		int x = range->GetRow();
		intrusive_ptr<CellRange> sourceRange = dynamic_pointer_cast<CellRange>(sheet->GetRange(range->GetRow(), 1, range->GetRow(), columnCount));
		intrusive_ptr<CellRange> destRange = dynamic_pointer_cast<CellRange>(newSheet->GetRange(k + 1, 1, k + 1, columnCount));
		sheet->Copy(sourceRange, destRange, true);
		k++;
	}
}
```

---

# spire.xls c++ data validation
## set data validation on separate sheet
```cpp
//This is the first sheet
intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

sheet1->GetRange(L"B10")->SetText(L"Here is a dataValidation example.");
//This is the second sheet
intrusive_ptr<Worksheet> sheet2 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));
//The property is to enable the data can be from different sheet.
sheet2->GetParentWorkbook()->SetAllow3DRangesInDataValidation(true);
sheet1->GetRange(L"B11")->GetDataValidation()->SetDataRange(dynamic_pointer_cast<CellRange>(sheet2->GetRange(L"A1:A7")));
```

---

# spire.xls cpp subtotal
## create subtotal in Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Select data range
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B18"));
//Subtotal selected data
sheet->Subtotal(range, 0, { 1 }, SubtotalTypes::Sum, true, false, true);
```

---

# spire.xls cpp richtext
## write rich text with different font styles to excel cell
```cpp
// Create different font styles
intrusive_ptr<ExcelFont> fontBold = workbook->CreateExcelFont();
fontBold->SetIsBold(true);

intrusive_ptr<ExcelFont> fontUnderline = workbook->CreateExcelFont();
fontUnderline->SetUnderline(FontUnderlineType::Single);

intrusive_ptr<ExcelFont> fontItalic = workbook->CreateExcelFont();
fontItalic->SetIsItalic(true);

intrusive_ptr<ExcelFont> fontColor = workbook->CreateExcelFont();
fontColor->SetKnownColor(ExcelColors::Green);

// Get RichText object from cell and apply different font styles
intrusive_ptr<RichText> richText = dynamic_pointer_cast<RichText>(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B11"))->GetRichText());
richText->SetText(L"Bold and underlined and italic and colored text.");
richText->SetFont(0, 3, fontBold);
richText->SetFont(9, 18, fontUnderline);
richText->SetFont(24, 29, fontItalic);
richText->SetFont(35, 41, fontColor);
```

---

# spire.xls cpp cell access
## demonstrates different ways to access cells in an Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Access cell by its name
intrusive_ptr<CellRange> range1 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"));
wstring s1 = range1->GetText();

//Access cell by index of row and column
intrusive_ptr<CellRange> range2 = dynamic_pointer_cast<CellRange>(sheet->GetRange(2, 1));
wstring s2 = range2->GetText();

//Access cell in cell collection
intrusive_ptr<XlsRange> range3 = sheet->GetCells()->GetItem(2);
wstring s3 = range3->GetText();
```

---

# spire.xls cpp multiple fonts in cell
## Apply multiple fonts to different parts of text within a single Excel cell
```cpp
//Create a font object in workbook, setting the font color, size and type.
intrusive_ptr<ExcelFont> font1 = workbook->CreateExcelFont();
font1->SetKnownColor(ExcelColors::LightBlue);
font1->SetIsBold(true);
font1->SetSize(10);

//Create another font object specifying its properties.
intrusive_ptr<ExcelFont> font2 = workbook->CreateExcelFont();
font2->SetKnownColor(ExcelColors::Red);
font2->SetIsBold(true);
font2->SetIsItalic(true);
font2->SetFontName(L"Times New Roman");
font2->SetSize(11);

//Write a RichText string to the cell 'A1', and set the font for it.
intrusive_ptr<RichText> richText = dynamic_pointer_cast<RichText>(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"H5"))->GetRichText());
richText->SetText(L"This document was created with Spire.XLS for .NET.");
richText->SetFont(0, 29, font1);
richText->SetFont(31, 48, font2);
```

---

# spire.xls cpp cells
## auto fit columns and rows based on cell value
```cpp
//Set value for B8
intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8"));
cell->SetText(L"Welcome to Spire.XLs!");

//Set the cell style
intrusive_ptr<CellStyle> style = dynamic_pointer_cast<CellStyle>(cell->GetStyle());
style->GetFont()->SetSize(16);
style->GetFont()->SetIsBold(true);

//Autofit column width and row height based on cell value
cell->AutoFitColumns();
cell->AutoFitRows();
```

---

# spire.xls cpp cell format
## convert text to number format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Convert text string format to number format
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2:D8"))->ConvertToNumber();
```

---

# spire.xls cpp cell format
## copy cell format from one column to another
```cpp
//Copy the cell format from column 2 and apply to cells of column 5.
int count = sheet->GetRows()->GetCount();
for (int i = 1; i < count + 1; i++)
{
    dynamic_pointer_cast<CellRange>(sheet->GetRange((L"E" + std::to_wstring(i)).c_str()))->SetStyle(dynamic_pointer_cast<CellStyle>(dynamic_pointer_cast<CellRange>(sheet->GetRange((L"B" + std::to_wstring(i)).c_str()))->GetStyle()));
}
```

---

# spire.xls cpp cells count
## count number of cells in worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the number of cells in the worksheet
int cellCount = sheet->GetCells()->GetCount();
```

---

# c++ cut cells to other position
## cut cells from one position to another in excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<CellRange> Ori = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5"));
intrusive_ptr<CellRange> Dest = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A26:C30"));

//Copy the range to other position
sheet->Copy(Ori, Dest, true, true, true);

//Remove all content in original cells
for (int i = 0; i < Ori->GetCells()->GetCount(); i++)
{
    intrusive_ptr<CellRange> cr = Ori->GetCells()->GetItem(i);
    cr->ClearAll();
}
```

---

# spire.xls cpp cells
## detect and unmerge merged cells
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the merged cell ranges in the first worksheet and put them into a CellRange array.
intrusive_ptr<Spire::Xls::IList<XlsRange>> range = sheet->GetMergedCells();

//Traverse through the array and unmerge the merged cells.
for (int i = 0; i < range->GetCount(); i++)
{
    intrusive_ptr<XlsRange> cell = range->GetItem(i);
    cell->UnMerge();
}

workbook->Dispose();
```

---

# spire.xls cpp duplicate range
## duplicate cell range in excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Copy data from source range to destination range and maintain the format.
sheet->Copy(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6:F6")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A16:F16")), true);
```

---

# c++ excel cell emptying
## demonstrates different methods to empty or clear cells in excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the value as null to remove the original content from the Excel Cell.
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C6"))->SetValue(L"");

//Clear the contents to remove the original content from the Excel Cell.
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B6"))->ClearContents();

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D6"))->ClearAll();
```

---

# Filter Cells by Cell Color
## Apply color filter to Excel cells based on cell color
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create an auto filter in the sheet and specify the range to be filterd
sheet->GetAutoFilters()->SetRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"G1:G19")));

//Get the coloumn to be filterd
intrusive_ptr<FilterColumn> filtercolumn = sheet->GetAutoFilters()->Get(0);

//Add a color filter to filter the column based on cell color
(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->AddFillColorFilter(filtercolumn, Spire::Xls::Color::GetRed());

//Filter the data.
(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->Filter();
```

---

# Finding Cells with Style Name
## This code demonstrates how to find all cells in a worksheet that have the same style name as a reference cell and mark them.
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the cell style name
wstring styleName = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->GetCellStyleName();

intrusive_ptr<CellRange> ranges = dynamic_pointer_cast<CellRange>(sheet->GetAllocatedRange());
for (int i = 0; i < ranges->GetCells()->GetCount(); i++)
{
    intrusive_ptr<CellRange> cr = ranges->GetCells()->GetItem(i);
    //Find the cells which have the same style name
    if (!wcscmp(cr->GetCellStyleName(), styleName.c_str()))
    {
        //Set value
        cr->SetValue(L"Same style");
    }
}
```

---

# spire.xls cpp find formula cells
## find cells containing specific formula in Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str()); // inputFile should contain the path to the Excel file

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Find the cells that contain formula "=SUM(A11,A12)"
intrusive_ptr<Spire::Xls::IList<CellRange>> ranges = sheet->FindAll(L"=SUM(A11,A12)", FindType::Formula, ExcelFindOptions::None);

//Create a string builder
std::wstring* builder = new std::wstring();

//Append the address of found cells to builder
if (ranges->GetCount() != 0)
{
    for (int i = 0; i < ranges->GetCount(); i++)
    {
        intrusive_ptr<CellRange> cr = ranges->GetItem(i);
        wstring address = cr->GetRangeAddress();
        builder->append(L"The address of found cell is: " + address);
    }
}
else
{
    builder->append(L"No cell contain the formula");
}
```

---

# spire.xls cpp cell address
## get cell address and range information
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

std::wstring* builder = new std::wstring();

//Get a cell range
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B5"));

//Get address of range
wstring address = range->GetRangeAddressLocal();
builder->append(L"Address of range: " + address);

//Get the cell count of range
int count = range->GetCellsCount();
builder->append(L"Cell count of range: " + std::to_wstring(count));

//Get the address of the entire column of range
wstring entireColAddress = range->GetEntireColumn()->GetRangeAddressLocal();
builder->append(L"Address of entire column of the range: " + entireColAddress);

//Get the address of the entire row of range
wstring entireRowAddress = range->GetEntireColumn()->GetRangeAddressLocal();
builder->append(L"Address of entire row of the range " + entireRowAddress);
```

---

# spire.xls cpp get cell displayed text
## Get the displayed text of a cell with formatting applied
```cpp
//Set value for B8
intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8"));
cell->SetNumberValue(0.012345);

//Set the cell style
intrusive_ptr<CellStyle> style = dynamic_pointer_cast<CellStyle>(cell->GetStyle());
style->SetNumberFormat(L"0.00");

//Get the cell value
wstring cellValue = cell->GetValue();

//Get the displayed text of the cell
wstring displayedText = cell->GetDisplayedText();
```

---

# c++ get cell value by cell name
## retrieve cell value from excel using cell name reference
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Specify a cell by its name.
intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"));

std::wstring* content = new std::wstring();

//Get vaule of cell L"A2".
wstring cellValue = cell->GetValue();
content->append(L"The vaule of cell A2 is: " + cellValue);
```

---

# spire.xls cpp ranges
## get intersection of two ranges
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the two ranges.
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:D7"))->Intersect(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:E8")));

//Get the intersection of the two ranges.
for (int i = 0; i < range->GetCells()->GetCount(); i++)
{
    intrusive_ptr<CellRange> cr = range->GetCells()->GetItem(i);
    content->append(cr->GetValue());
}
```

---

# spire.xls cpp hide cell content
## hide cell content by setting number format
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Hide the area by setting the number format as L";;;".
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5:D6"))->SetNumberFormat(L";;;");
```

---

# spire.xls cpp merge cells
## merge cells in excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Merge the seventh column in Excel file.
sheet->GetColumns()->GetItem(6)->Merge();

//Merge the particular range in Excel file.
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A14:D14"))->XlsRange::Merge();
```

---

# Copy Formula Values Only
## Copy only formula values from one cell range to another in Excel
```cpp
//Set the copy option
CopyRangeOptions copyOptions = CopyRangeOptions::OnlyCopyFormulaValue;

intrusive_ptr<CellRange> sourceRange = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6:E6"));
sheet->Copy(sourceRange, dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:E8")), copyOptions);

sourceRange->Copy(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A10:E10")), copyOptions);
```

---

# spire.xls cpp cell formatting
## set cell fill pattern and color
```cpp
//Set cell color
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B7:F7"))->GetStyle()->SetColor(Spire::Xls::Color::GetYellow());
//Set cell fill pattern
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8:F8"))->GetStyle()->SetFillPattern(ExcelPatternType::Percent125Gray);
```

---

# spire.xls cpp cell formatting
## set DB number formatting for Excel cells
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();
workbook->CreateEmptySheets(1);

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set value for cells
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetNumberValue(123);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetNumberValue(456);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetNumberValue(789);

//Get the cell range
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:A3"));

//Set the DB num format
range->SetNumberFormat(L"[DBNum2][$-804]General");

//Auto fit columns
range->AutoFitColumns();
```

---

# spire.xls cpp cell formatting
## shrink text to fit in a cell
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//The cell range to shrink text.
intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B13:C13"));

//Enable ShrinkToFit.
intrusive_ptr<CellStyle> style = dynamic_pointer_cast<CellStyle>(cell->GetStyle());
style->SetShrinkToFit(true);
```

---

# Excel Cell Value Traversal
## Traverse through all cells in a worksheet and get their values
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the cell range collection 
intrusive_ptr<Spire::Xls::IList<XlsRange>> cellRangeCollection = sheet->GetCells();

//Traverse cells value
for (int i = 0; i < cellRangeCollection->GetCount(); i++)
{
    intrusive_ptr<XlsRange> cr = cellRangeCollection->GetItem(i);
    //Set string format for displaying
    wstring cell = cr->GetRangeAddress();
    wstring result = L"Cell: " + cell + L"   Value: " + cr->GetValue();
}
```

---

# c++ ungroup excel cells
## ungroup rows in excel worksheet
```cpp
//Ungroup the row 10 to 12.
sheet->UngroupByRows(10, 12);

//Ungroup the row 16 to 19.
sheet->UngroupByRows(16, 19);
```

---

# spire xls cpp unmerge cells
## unmerge specific cells in an excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Unmerge the cells.
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"F2"))->UnMerge();

//Unmerge the cells.
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"F7"))->UnMerge();
```

---

# spire.xls cpp explicit line breaks
## use explicit line breaks in excel cells
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Specify a cell range
intrusive_ptr<CellRange> c5 = dynamic_pointer_cast<CellRange>(sheet1->GetRange(L"C5"));

//Set the cell width for specified range
sheet1->SetColumnWidth(c5->GetColumn(), 70);

//Put the string value with explicit line breaks
c5->SetValue(L"Spire.XLS for C++ is a professional Excel C++ API\n that can be used to create, read, write and convert Excel files in any type of C++ application.\n Spire.XLS for C++ offers object model Excel API for speeding up Excel programming\n in C++ platform -create new Excel documents from template, edit existing Excel documents and convert Excel files.");

//Set Text wrap
c5->SetIsWrapText(true);
```

---

# spire.xls cpp text wrapping
## wrap or unwrap text in excel cells
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Wrap the excel text
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->GetStyle()->SetWrapText(true);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D1"))->GetStyle()->SetWrapText(true);

//Unwrap the excel text
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->GetStyle()->SetWrapText(false);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2"))->GetStyle()->SetWrapText(false);
```

---

# AutoFit Column in Range
## Demonstrates how to auto-fit a specific column within a range in an Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Auto-fit the column of the worksheet
sheet->AutoFitColumn(2, 2, 5);
```

---

# spire.xls cpp autofit row
## auto fit row in specified range
```cpp
//Autofit the second row of the worksheet
sheet->AutoFitRow(2, 1, 2, false);
```

---

# Check AutoFit Row or Column
## This code demonstrates how to check if a row or column is set to auto-fit in an Excel worksheet.
```cpp
// Check if the second row is auto fit
bool isRowAutofit = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetRowIsAutoFit(2);

// Check if the second column is auto fit
bool isColAutofit = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetColumnIsAutoFit(2);
```

---

# spire.xls cpp row column visibility
## check if row or column is hidden in Excel worksheet
```cpp
// Create a new Workbook object using intrusive_ptr smart pointer.
intrusive_ptr<Workbook> workbook = new Workbook();

// Retrieve the first worksheet from the workbook using dynamic_pointer_cast.
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Set the row and column indices for checking visibility.
int rowIndex = 2;
int columnIndex = 2;

// Check if the specified row is hidden.
bool rowIsHide = sheet->GetRowIsHide(rowIndex);

// Check if the specified column is hidden.
bool columnIsHide = sheet->GetColumnIsHide(columnIndex);

// Dispose of the workbook object.
workbook->Dispose();
```

---

# spire.xls cpp copy range with options
## copy cell range from one worksheet to another with style preservation and reference update
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add a new worksheet as destination sheet
intrusive_ptr<Worksheet> destinationSheet = workbook->GetWorksheets()->Add(L"DestSheet");

//Specify a copy range of original sheet
intrusive_ptr<CellRange> cellRange = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D4"));

//Copy the specified range to added worksheet and keep original styles and update reference
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->Copy(cellRange, dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1)), 2, 1, true, true);
```

---

# Spire.XLS C++ Delete Blank Rows and Columns
## Remove empty rows and columns from an Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Delete blank rows from the worksheet.
for (int i = sheet->GetRows()->GetCount() - 1; i >= 0; i--)
{
    if (sheet->GetRows()->GetItem(i)->GetIsBlank())
    {
        sheet->DeleteRow(i + 1);
    }
}

//Delete blank columns from the worksheet.
for (int j = sheet->GetColumns()->GetCount() - 1; j >= 0; j--)
{
    if (sheet->GetColumns()->GetItem(j)->GetIsBlank())
    {
        sheet->DeleteColumn(j + 1);
    }
}
```

---

# spire.xls cpp delete rows columns
## delete multiple rows and columns from excel worksheet
```cpp
//Delete 4 rows from the fifth row
sheet->DeleteRow(5, 4);

//Delete 2 columns from the second column
sheet->DeleteColumn(2, 2);
```

---

# spire.xls cpp get default row column count
## Get the default row and column count of an Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Clear all worksheets
workbook->GetWorksheets()->Clear();

//Create a new worksheet
intrusive_ptr<Worksheet> sheet = workbook->CreateEmptySheet();

//Get row and column count
int rowCount = sheet->GetRows()->GetCount();
int columnCount = sheet->GetColumns()->GetCount();
```

---

# spire.xls cpp rows columns
## Group rows and columns in Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Group rows
sheet->GroupByRows(1, 5, false);
// Group columns
sheet->GroupByColumns(1, 3, false);
```

---

# c++ excel row column headers
## hide or show row and column headers in excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Hide the headers of rows and columns
sheet->SetRowColumnHeadersVisible(false);

//Show the headers of rows and columns
//sheet->SetRowColumnHeadersVisible(true);
```

---

# spire.xls cpp hide rows columns
## hide specific rows and columns in Excel worksheet
```cpp
// Hiding the column of the worksheet
sheet->HideColumn(2);
//Hiding the row of the worksheet
sheet->HideRow(4);
```

---

# spire.xls cpp rows and columns
## insert rows and columns in Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Inserting a row into the worksheet 
sheet->InsertRow(2);
//Inserting a column into the worksheet 
sheet->InsertColumn(2);
//Inserting multiple rows into the worksheet
sheet->InsertRow(5, 2);
//Inserting multiple columns into the worksheet
sheet->InsertColumn(5, 2);
```

---

# spire.xls cpp remove row
## remove row based on keyword
```cpp
//Find the string
intrusive_ptr<CellRange> cr = sheet->FindString(L"Address", false, false);

//Delete the row which includes the string
sheet->DeleteRow(cr->GetRow());
```

---

# spire.xls cpp column width
## set column width in pixels
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the width of the third column to 400 pixels
sheet->SetColumnWidthInPixels(3, 400);
```

---

# c++ excel default column width
## Set default column width for Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set default column width
sheet->SetDefaultColumnWidth(25);
```

---

# spire.xls cpp row column style
## set default style for rows and columns
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create a cell style and set the color
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"Mystyle");
style->SetColor(Spire::Xls::Color::GetYellow());

//Set the default style for the first row and column 
sheet->SetDefaultRowStyle(1, style);
sheet->SetDefaultColumnStyle(1, style);
```

---

# spire.xls cpp row height
## set default row height for worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set default row height
sheet->SetDefaultRowHeight(30);
```

---

# spire.xls cpp row column
## Set row height and column width in Excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Setting the width to 30
sheet->SetColumnWidth(4, 30);
// Setting the height to 30
sheet->SetRowHeight(4, 30);
```

---

# c++ excel summary column direction
## set summary column direction in excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Group Columns
sheet->GroupByColumns(1, 4, true);

//Set summary columns to right of details
sheet->GetPageSetup()->SetIsSummaryRowBelow(true);
```

---

# c++ set summary row direction
## set summary row position in excel worksheet
```cpp
//Group rows
sheet->GroupByRows(1, 4, true);
//Set summary rows above details
sheet->GetPageSetup()->SetIsSummaryRowBelow(false);
```

---

# spire.xls cpp rows columns
## unhide specific row and column in Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Unhide the row
sheet->ShowRow(15);

//Unhide the column
sheet->ShowColumn(4);
```

---

# spire.xls cpp image alignment
## align picture within a cell in Excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

sheet->GetRange(L"A1")->SetText(L"Align Picture Within A Cell:");
sheet->GetRange(L"A1")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Top);
intrusive_ptr<ExcelPicture> picture = ExcelPicture::Dynamic_cast<ExcelPicture>(sheet->GetPictures()->Add(1, 1, L"image_path"));

//Adjust the column width and row height so that the cell can contain the picture.
sheet->GetColumns()->GetItem(0)->SetColumnWidth(40);
sheet->GetRows()->GetItem(0)->SetRowHeight(200);

//Vertically and horizontally align the image.
picture->SetLeftColumnOffset(100);
picture->SetTopRowOffset(25);
```

---

# c++ excel picture copy
## copy picture from one worksheet to another in excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add a new worksheet as destination sheet
intrusive_ptr<Worksheet> destinationSheet = workbook->GetWorksheets()->Add(L"DestSheet");
//Get the first picture from the first worksheet
intrusive_ptr<XlsBitmapShape> sourcePicture = sheet->GetPictures()->Get(0);
//Get the image
intrusive_ptr<Stream> image = sourcePicture->GetPicture();
//Add the image into the added worksheet 
destinationSheet->GetPictures()->Add(2, 2, image);
```

---

# spire.xls cpp images
## delete all images from worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Delete all images from the worksheet.
for (int i = sheet->GetPictures()->GetCount() - 1; i >= 0; i--)
{
	sheet->GetPictures()->Get(i)->XlsShape::Remove();
}
```

---

# spire xls cpp image crop position
## get crop position of an image in excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the image from the first sheet
intrusive_ptr<XlsBitmapShape> picture = sheet->GetPictures()->Get(0);

//Get the cropped position
int left = picture->GetLeft();
int top = picture->GetTop();
int width = picture->GetWidth();
int height = picture->GetHeight();

//Set string format for displaying
wstring displayString = L"Crop position: Left " + std::to_wstring(left) + L"\r\nCrop position: Top " + std::to_wstring(top) + L"\r\nCrop position: Width " + std::to_wstring(width) + L"\r\nCrop position: Height " + std::to_wstring(height);
```

---

# spire.xls cpp background image
## insert background image to Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Open an image. 
intrusive_ptr<Stream> im = new Stream(inputImage.c_str());
intrusive_ptr<Stream> bm = Object::Convert<Stream>(im);
//Set the image to be background image of the worksheet.
sheet->GetPageSetup()->SetBackgoundImage(bm);
```

---

# spire.xls cpp image positioning
## locate and position images in excel worksheet
```cpp
intrusive_ptr<XlsBitmapShape> pic = sheet->GetPictures()->Get(0);
pic->SetLeftColumnOffset(300);
pic->SetTopRowOffset(300);
```

---

# spire.xls cpp image
## set picture offset in excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Insert a picture
intrusive_ptr<ExcelPicture> pic = ExcelPicture::Dynamic_cast<ExcelPicture>(sheet->GetPictures()->Add(2, 2, "image_path"));

//Set left offset and top offset from the current range
pic->SetLeftColumnOffset(200);
pic->SetTopRowOffset(100);
```

---

# Excel Picture Reference Range
## Setting reference range for a picture in Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Spire.XLS");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetValue(L"E-iceblue");

//Get the first picture in worksheet
intrusive_ptr<XlsBitmapShape> picture = sheet->GetPictures()->Get(0);

//Set the reference range of the picture to A1:B3
picture->SetRefRange(L"A1:B3");
```

---

# spire.xls cpp image
## extract image from excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFilePath);

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first image
intrusive_ptr<XlsBitmapShape> pic = sheet->GetPictures()->Get(0);
pic->GetPicture()->Save(outputFilePath);
workbook->Dispose();
```

---

# spire.xls cpp image
## remove picture border in excel
```cpp
//Get the first picture from the first worksheet
intrusive_ptr<XlsBitmapShape> picture = sheet->GetPictures()->Get(0);

//Remove the picture border
//Method-1:
picture->GetLine()->SetVisible(false);

//Method-2:
//picture->GetLine()->SetWeight(0);
```

---

# spire.xls cpp image manipulation
## reset size and position for image in excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add a picture to the first worksheet.
intrusive_ptr<IPictureShape> picture = dynamic_pointer_cast<XlsPicturesCollection>(sheet->GetPictures())->Add(1, 1, inputFile.c_str());

//Set the size for the picture.
picture->SetWidth(200);
picture->SetHeight(200);

//Set the position for the picture.
picture->SetLeft(200);
picture->SetTop(100);
```

---

# spire.xls cpp chart image
## set image offset of chart
```cpp
//Add chart and background image to sheet as comparison
intrusive_ptr<Chart> chart1 = sheet1->GetCharts()->Add(ExcelChartType::ColumnClustered);
chart1->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D1:E8")));
chart1->SetSeriesDataFromRange(false);

//Chart Position
chart1->SetLeftColumn(1);
chart1->SetTopRow(11);
chart1->SetRightColumn(8);
chart1->SetBottomRow(33);

//Add picture as background
chart1->GetChartArea()->GetFill()->CustomPicture(new Stream(imagePath), L"None");
chart1->GetChartArea()->GetFill()->SetTile(false);

//Set the image offset
chart1->GetChartArea()->GetFill()->GetPicStretch()->SetLeft(20);
chart1->GetChartArea()->GetFill()->GetPicStretch()->SetTop(20);
chart1->GetChartArea()->GetFill()->GetPicStretch()->SetRight(5);
chart1->GetChartArea()->GetFill()->GetPicStretch()->SetBottom(5);
```

---

# spire.xls cpp images
## write image to excel worksheet
```cpp
using namespace Spire::Xls;

//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add an image to the specific cell at row 14, column 5
dynamic_pointer_cast<XlsPicturesCollection>(sheet->GetPictures())->Add(14, 5, imagePath);

workbook->Dispose();
```

---

# spire.xls cpp comment
## add comment with author to Excel cell
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the range that will add comment
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"));

//Set the author and comment content
wstring author = L"E-iceblue";
wstring text = L"This is demo to show how to add a comment with editable Author property.";

//Add comment to the range and set properties
intrusive_ptr<ExcelComment> comment = range->AddExcelComment();
comment->SetWidth(200);
comment->SetVisible(true);
comment->SetText(author.empty() ? text.c_str() : (author + L":\n" + text).c_str());

//Set the font of the author
intrusive_ptr<IFont> font = range->GetWorksheet()->GetWorkbook()->CreateFont();
font->SetFontName(L"Tahoma");
font->SetKnownColor(ExcelColors::Black);
font->SetIsBold(true);
comment->GetRichText()->SetFont(0, author.length(), font);
```

---

# spire.xls cpp comment with picture
## Add Excel comment with picture
```cpp
//Add the comment
intrusive_ptr<ExcelComment> comment = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C6"))->AddExcelComment();
//Load the image file
intrusive_ptr<Stream> image = new Stream(inputFile.c_str());

comment->GetFill()->CustomPicture(image, L"logo.png");
//Set the height and width of comment
comment->SetVisible(true);
```

---

# spire.xls cpp comment
## edit Excel comment text
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first comment.
intrusive_ptr<XlsComment> comment = sheet->GetComments()->Get(0);

//Edit the comment.
comment->SetText(L"This comment has been edited by Spire.XLS.");
```

---

# spire.xls cpp comment
## hide or show comments in excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Hide the second comment
sheet->GetComments()->Get(1)->SetIsVisible(false);

//Show the third comment
sheet->GetComments()->Get(2)->SetIsVisible(true);
```

---

# spire.xls cpp comment
## modify and remove comments from Excel worksheet
```cpp
//Get all comments from the first worksheet
intrusive_ptr<XlsCommentsCollection> comments = dynamic_pointer_cast<XlsCommentsCollection>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetComments());
//Change the content of the first comment
comments->Get(0)->SetText(L"This comment has been changed.");
//Remove the second comment
comments->Get(1)->Remove();
```

---

# spire.xls cpp comment fill color
## Set comment fill color in Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create Excel font
intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
font->SetFontName(L"Arial");
font->SetSize(11);
font->SetKnownColor(ExcelColors::Orange);

//Add the comment
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"));
range->GetComment()->SetText(L"This is a comment");
wstring text = range->GetComment()->GetText();
range->GetComment()->GetRichText()->SetFont(0, (text.size() - 1), font);

//Set comment Color
range->GetComment()->GetFill()->SetFillType(ShapeFillType::SolidColor);
range->GetComment()->GetFill()->SetForeColor(Spire::Xls::Color::GetSkyBlue());

range->GetComment()->SetVisible(true);
```

---

# spire.xls c++ comment text rotation
## set text rotation for excel cell comment
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create Excel font
intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
font->SetFontName(L"Arial");
font->SetSize(11);
font->SetKnownColor(ExcelColors::Orange);

//Add the comment
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E1"));
range->GetComment()->SetText(L"This is a comment");
wstring text = range->GetComment()->GetText();
range->GetComment()->GetRichText()->SetFont(0, (text.size() - 1), font);

// Set its vertical and horizontal alignment 
range->GetComment()->SetVAlignment(CommentVAlignType::Center);
range->GetComment()->SetHAlignment(CommentHAlignType::Right);

//Set the comment text rotation
range->GetComment()->SetTextRotation(TextRotationType::LeftToRight);
```

---

# spire.xls cpp comment position alignment
## Set position and alignment of Excel comments
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set two font styles which will be used in comments
intrusive_ptr<ExcelFont> font1 = workbook->CreateExcelFont();
font1->SetFontName(L"Calibri");
font1->SetColor(Spire::Xls::Color::GetFirebrick());
font1->SetIsBold(true);
font1->SetSize(12);
intrusive_ptr<ExcelFont> font2 = workbook->CreateExcelFont();
font2->SetFontName(L"Calibri");
font2->SetColor(Spire::Xls::Color::GetBlue());
font2->SetSize(12);
font2->SetIsBold(true);

//Add comment 1 and set its size, text, position and alignment
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"G5"))->SetText(L"Spire.XLS");
intrusive_ptr<IComment> Comment1 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"G5"))->GetComment();
Comment1->SetIsVisible(true);
Comment1->SetHeight(150);
Comment1->SetWidth(300);
Comment1->GetRichText()->SetText(L"Spire.XLS for .Net:\nStandalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. ");
Comment1->GetRichText()->SetFont(0, 19, font1);
Comment1->SetTextRotation(TextRotationType::LeftToRight);
//Set the position of Comment
Comment1->SetTop(20);
Comment1->SetLeft(40);
//Set the alignment of text in Comment
Comment1->SetVAlignment(CommentVAlignType::Center);
Comment1->SetHAlignment(CommentHAlignType::Justified);

//Add comment2 and set its size, text, position and alignment for comparison
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D14"))->SetText(L"E-iceblue");
intrusive_ptr<IComment> Comment2 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D14"))->GetComment();
Comment2->SetIsVisible(true);
Comment2->SetHeight(150);
Comment2->SetWidth(300);
Comment2->GetRichText()->SetText(L"About E-iceblue: \nWe focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents.");
Comment2->SetTextRotation(TextRotationType::LeftToRight);
Comment2->GetRichText()->SetFont(0, 16, font2);
//Set the position of Comment
Comment2->SetTop(170);
Comment2->SetLeft(450);
//Set the alignment of text in Comment
Comment2->SetVAlignment(CommentVAlignType::Top);
Comment2->SetHAlignment(CommentHAlignType::Justified);
```

---

# spire.xls cpp comments
## write regular and rich text comments in excel cells
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Creates font
intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
font->SetFontName(L"Arial");
font->SetSize(11);
font->SetKnownColor(ExcelColors::Orange);
intrusive_ptr<ExcelFont> fontBlue = workbook->CreateExcelFont();
fontBlue->SetKnownColor(ExcelColors::LightBlue);
intrusive_ptr<ExcelFont> fontGreen = workbook->CreateExcelFont();
fontGreen->SetKnownColor(ExcelColors::LightGreen);

intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B11"));
range->SetText(L"Regular comment");
range->GetComment()->SetText(L"Regular comment");
range->AutoFitColumns();
//Regular comment

range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B12"));
range->SetText(L"Rich text comment");
range->GetRichText()->SetFont(0, 16, font);
range->AutoFitColumns();
//Rich text comment
range->GetComment()->GetRichText()->SetText(L"Rich text comment");
range->GetComment()->GetRichText()->SetFont(0, 4, fontGreen);
range->GetComment()->GetRichText()->SetFont(5, 9, fontBlue);
```

---

# CSV to Excel Conversion
## Convert CSV file to Excel format with error handling and column auto-fitting
```cpp
// Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

// Load the CSV document from disk
workbook->LoadFromFile(L"CSVToExcel.csv", L",");

// Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Set ignore error options for specific range
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2:E19"))->SetIgnoreErrorOptions(IgnoreErrorType::NumberAsText);
sheet->GetAllocatedRange()->AutoFitColumns();

// Save to Excel file
workbook->SaveToFile(L"CSVToExcel_out.xlsx", ExcelVersion::Version2013);
workbook->Dispose();
```

---

# spire.xls cpp conversion
## convert CSV file to PDF format
```cpp
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str(), L",");

//Set the SheetFitToPage property as true
workbook->GetConverterSetting()->SetSheetFitToPage(true);

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Autofit a column if the characters in the column exceed column width
for (int i = 1; i < sheet->GetColumns()->GetCount(); i++)
{
    sheet->AutoFitColumn(i);
}

//Save to file.
workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
workbook->Dispose();
```

---

# Excel to PDF Conversion
## Convert each worksheet to a separate PDF file
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
{
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
    wstring FileName = outputFile + sheet->GetName() + L".pdf";
    //Save the sheet to PDF
    sheet->SaveToPdf(FileName.c_str());
}
```

---

# Excel to PDF Conversion with Width Fitting
## Set worksheet to fit width when converting Excel to PDF
```cpp
// Iterate through all worksheets
for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
{
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
    // Auto fit page height
    sheet->GetPageSetup()->SetFitToPagesTall(0);
    // Fit one page width
    sheet->GetPageSetup()->SetFitToPagesWide(1);
}
```

---

# html to excel conversion
## convert HTML file to Excel format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Save to file.
workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
workbook->Dispose();
```

---

# spire.xls cpp et file conversion
## load and save ET format files
```cpp
// Create a new Workbook object using intrusive_ptr smart pointer.
intrusive_ptr<Workbook> workbook = new Workbook();

// Load the workbook from the specified input file.
workbook->LoadFromFile(inputFile.c_str());

// Save the workbook to the specified output file with the specified file format (FileFormat::ET).
workbook->SaveToFile(outputFile.c_str(), FileFormat::ET);

// Dispose of the workbook object.
workbook->Dispose();
```

---

# Office Open XML to Excel Conversion
## Convert Office Open XML format to Excel format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load XML file
ifstream fs(inputFile.c_str(), ios::in | ios::binary);
intrusive_ptr<Stream> fileStream = new Stream(fs);
workbook->LoadFromXml(fileStream);

//Save to Excel file
workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
workbook->Dispose();
```

---

# Excel Range to PDF Conversion
## Convert a selected range from an Excel worksheet to PDF format
```cpp
using namespace Spire::Xls;

wstring inputFile = L"ConversionSample1.xlsx";
wstring outputFile = L"SelectedRangeToPDF.pdf";

//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add a new sheet to workbook
workbook->GetWorksheets()->Add(L"newsheet");
//Copy your area to new sheet.
dynamic_pointer_cast<CellRange>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetRange(L"A9:E15"))->Copy(dynamic_pointer_cast<CellRange>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetRange(L"A9:E15")), false, true);
//Auto fit column width
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetRange(L"A9:E15")->AutoFitColumns();

//Save the selected range to PDF
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->SaveToPdf(outputFile.c_str());
workbook->Dispose();
```

---

# Spire.XLS C++ Sheet to Image Conversion
## Convert Excel worksheet to image
```cpp
using namespace Spire::Xls;

int main() {
    //Create a workbook
    intrusive_ptr<Workbook> workbook = new Workbook();

    //Get the first worksheet
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

    //Convert sheet to image
    auto image = sheet->ToImage(sheet->GetFirstRow(), sheet->GetFirstColumn(), sheet->GetLastRow(), sheet->GetLastColumn());
    
    workbook->Dispose();
}
```

---

# spire.xls cpp conversion
## convert specific cell ranges to different image formats
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Specify Cell Ranges and Save to certain Image formats
sheet->ToImage(1, 1, 7, 5)->Save((outputFile + L"SpecificCellsToImage.png").c_str());
sheet->ToImage(8, 1, 15, 5)->Save((outputFile + L"SpecificCellsToImage.jpg").c_str());
sheet->ToImage(17, 1, 23, 5)->Save((outputFile + L"SpecificCellsToImage.bmp").c_str());
```

---

# spire.xls cpp font directory
## specify custom font directory for Excel to PDF conversion
```cpp
// Create a new Workbook object using intrusive_ptr smart pointer.
intrusive_ptr<Workbook> workbook = new Workbook();

// Load the workbook from the specified input file.
workbook->LoadFromFile(inputFile.c_str());

// Create a vector to store custom font file paths as LPCWSTR_S (wide string) elements.
vector<LPCWSTR_S> fonts;

// Add the inputFontFile path to the fonts vector.
fonts.push_back(inputFontFile.c_str());

// Set the custom font file directory for the workbook using the fonts vector.
workbook->SetCustomFontFileDirectory(fonts);

// Save the workbook to the specified output file in PDF format.
workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);

// Dispose of the workbook object.
workbook->Dispose();
```

---

# spire.xls cpp conversion
## convert Excel worksheet to CSV format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//convert to CSV file
sheet->SaveToFile(outputFile.c_str(), L",", Encoding::GetUTF8());
workbook->Dispose();
```

---

# Excel to CSV Conversion
## Convert Excel file to CSV with filtered values
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Convert to CSV file with filtered value
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->SaveToFile(outputFile.c_str(), L";", false);
workbook->Dispose();
```

---

# Spire.XLS C++ Excel to HTML Conversion
## Convert Excel worksheet to HTML format with embedded images
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(/* path to input Excel file */);

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<HTMLOptions> options = new HTMLOptions();
options->SetImageEmbedded(true);

//Save to HTML file
sheet->SaveToHtml(/* path to output HTML file */);
workbook->Dispose();
```

---

# spire.xls cpp conversion
## convert excel worksheet to html stream
```cpp
//Set the html options
intrusive_ptr<HTMLOptions> options = new HTMLOptions();
options->SetImageEmbedded(true);
//Save sheet to html stream
intrusive_ptr<Stream> stream = new Stream();

//Save worksheet to html stream
sheet->SaveToHtml(stream, options);
```

---

# spire.xls cpp conversion
## convert excel worksheet to image without white space
```cpp
//Set the margin as 0 to remove the white space around the image
sheet->GetPageSetup()->SetLeftMargin(0);
sheet->GetPageSetup()->SetBottomMargin(0);
sheet->GetPageSetup()->SetTopMargin(0);
sheet->GetPageSetup()->SetRightMargin(0);
intrusive_ptr<Stream> image = sheet->ToImage(sheet->GetFirstRow(), sheet->GetFirstColumn(), sheet->GetLastRow(), sheet->GetLastColumn());
```

---

# c++ excel to ods conversion
## convert excel file to ods format
```cpp
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToODS.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToODS.ods";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//convert to ODS file
	workbook->SaveToFile(outputFile.c_str(), FileFormat::ODS);
	workbook->Dispose();
}
```

---

# Excel to OFD Conversion
## Convert Excel file to OFD format using Spire.XLS
```cpp
// Create a new Workbook object
intrusive_ptr<Workbook> workbook = new Workbook();

// Load the workbook from the input file
workbook->LoadFromFile(inputFile.c_str());

// Save the workbook to OFD format
workbook->SaveToFile(outputFile.c_str(), FileFormat::OFD);

// Dispose the workbook object
workbook->Dispose();
```

---

# spire.xls cpp conversion
## convert excel workbook to office open xml format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Save to Office Open XML format
workbook->SaveAsXml(outputFile.c_str());
```

---

# Convert Excel to PDF
## This code demonstrates how to convert an Excel file to PDF format using Spire.XLS for C++
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

workbook->GetConverterSetting()->SetSheetFitToPage(true);

//Save to file.
workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
workbook->Dispose();
```

---

# excel to pdf conversion
## convert excel file to pdf format using Spire.XLS library
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Convert excel to pdf
workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
workbook->Dispose();
```

---

# Excel to PDF Conversion with Page Size Change
## Convert Excel file to PDF while changing page size to A3
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document
workbook->LoadFromFile(inputFile.c_str());

//Change page size for all worksheets
for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
{
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
    //Change the page size
    sheet->GetPageSetup()->SetPaperSize(PaperSizeType::PaperA3);
}

//Save to PDF format
workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
workbook->Dispose();
```

---

# spire.xls cpp conversion
## convert Excel to PostScript format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Save to file.
workbook->SaveToFile(outputFile.c_str(), FileFormat::PostScript);
workbook->Dispose();
```

---

# spire.xls cpp conversion
## convert excel worksheets to svg format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
{
    intrusive_ptr<Stream> fileStream = new Stream();
    dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i))->ToSVGStream(fileStream, 0, 0, 0, 0);
    fileStream->Save((outputFile + L"sheet-" + std::to_wstring(i) + L".svg").c_str());
}

workbook->Dispose();
```

---

# Excel to Text Conversion
## Convert Excel worksheet to text file format
```cpp
// Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

// Load the Excel document from disk
workbook->LoadFromFile("ConversionSample1.xlsx");

// Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Save to text file
sheet->SaveToFile("ExceltoTxt.txt", L" ", Encoding::GetUTF8());
workbook->Dispose();
```

---

# Excel to TIFF Conversion
## Convert Excel file to TIFF image format
```cpp
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Save to TIFF file
workbook->SaveToFile(outputFile.c_str());
workbook->Dispose();
```

---

# c++ excel to xps conversion
## convert excel file to xps format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Save to XPS file
workbook->SaveToFile(outputFile.c_str(), Spire::Xls::FileFormat::XPS);
workbook->Dispose();
```

---

# Excel to HTML Conversion
## Convert Excel workbook to HTML format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Save to HTML file
workbook->SaveToHtml(outputFile.c_str());
workbook->Dispose();
```

---

# spire.xls cpp conversion
## convert XLS to XLSM format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile("input_file_path.xls");

//Save to file.
workbook->SaveToFile("output_file_path.xlsm", ExcelVersion::Version2007);
workbook->Dispose();
```

---

# spire.xls cpp autofilter
## filter blank cells in Excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Match the blank data
(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->MatchBlanks(0);

//Filter
(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->Filter();
```

---

# spire.xls cpp autofilter
## apply autofilter to show non-blank cells
```cpp
//Match the non blank data
(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->MatchNonBlanks(0);

//Filter
(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->Filter();
```

---

# spire.xls cpp filter
## create auto filter in Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create filter
sheet->GetAutoFilters()->SetRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:J1")));
```

---

# spire.xls cpp data validation
## create different types of data validation in Excel cells
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Decimal DataValidation
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B11"))->SetText(L"Input Number(3-6):");
intrusive_ptr<CellRange> rangeNumber = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B12"));
//Set the operator for the data validation.
rangeNumber->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);
//Set the value or expression associated with the data validation.
rangeNumber->GetDataValidation()->SetFormula1(L"3");
//The value or expression associated with the second part of the data validation.
rangeNumber->GetDataValidation()->SetFormula2(L"6");
//Set the data validation type.
rangeNumber->GetDataValidation()->SetAllowType(CellDataType::Decimal);
//Set the data validation error message.
rangeNumber->GetDataValidation()->SetErrorMessage(L"Please input correct number!");
//Enable the error.
rangeNumber->GetDataValidation()->SetShowError(true);
rangeNumber->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

//Date DataValidation
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B14"))->SetText(L"Input Date:");
intrusive_ptr<CellRange> rangeDate = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B15"));
rangeDate->GetDataValidation()->SetAllowType(CellDataType::Date);
rangeDate->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);
rangeDate->GetDataValidation()->SetFormula1(L"1/1/1970");
rangeDate->GetDataValidation()->SetFormula2(L"12/31/1970");
rangeDate->GetDataValidation()->SetErrorMessage(L"Please input correct date!");
rangeDate->GetDataValidation()->SetShowError(true);
rangeDate->GetDataValidation()->SetAlertStyle(AlertStyleType::Warning);
rangeDate->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

//TextLength DataValidation
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B17"))->SetText(L"Input Text:");
intrusive_ptr<CellRange> rangeTextLength = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B18"));
rangeTextLength->GetDataValidation()->SetAllowType(CellDataType::TextLength);
rangeTextLength->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::LessOrEqual);
rangeTextLength->GetDataValidation()->SetFormula1(L"5");
rangeTextLength->GetDataValidation()->SetErrorMessage(L"Enter a Valid String!");
rangeTextLength->GetDataValidation()->SetShowError(true);
rangeTextLength->GetDataValidation()->SetAlertStyle(AlertStyleType::Stop);
rangeTextLength->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

sheet->AutoFitColumn(2);
```

---

# spire.xls cpp filtering
## filter cells by string using custom criteria
```cpp
// Retrieve the first worksheet from the workbook using dynamic_pointer_cast.
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Get the AutoFiltersCollection from the worksheet.
intrusive_ptr<AutoFiltersCollection> filters = dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters());

// Set the range for filtering to column D (D1:D19).
filters->SetRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D1:D19")));

// Retrieve the first filter column from the AutoFiltersCollection.
intrusive_ptr<FilterColumn> filtercolumn = dynamic_pointer_cast<FilterColumn>(filters->Get(0));

// Create a SpireString object for the custom filter criteria ("South*").
intrusive_ptr<SpireString> cri = new SpireString(L"South*");

// Apply the custom filter to the filter column using Equal operator and the criteria.
filters->CustomFilter(filtercolumn, FilterOperatorType::Equal, cri);

// Apply the filters to filter the data.
filters->Filter();
```

---

# spire.xls cpp list data validation
## create list data validation in Excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set text for cells (list data)
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetText(L"Beijing");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8"))->SetText(L"New York");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A9"))->SetText(L"Denver");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A10"))->SetText(L"Paris");

//Set data validation for cell
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D10"));
range->GetDataValidation()->SetShowError(true);
range->GetDataValidation()->SetAlertStyle(AlertStyleType::Stop);
range->GetDataValidation()->SetErrorTitle(L"Error");
range->GetDataValidation()->SetErrorMessage(L"Please select a city from the list");
range->GetDataValidation()->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7:A10")));
```

---

# spire.xls cpp filter
## remove auto filters from Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Remove the auto filters.
(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->Clear();
```

---

# spire.xls cpp data validation
## remove data validation from excel ranges
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Create an array of rectangles, which is used to locate the ranges in worksheet.
std::vector<intrusive_ptr<Spire::Xls::Rectangle>> rectangles(1);

//Assign value to the first element of the array. This rectangle specifies the cells from A1 to B3.
rectangles[0] = Spire::Xls::Rectangle::FromLTRB(0, 0, 1, 2);

//Remove validations in the ranges represented by rectangles.
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetDVTable()->Remove(rectangles);
```

---

# spire.xls cpp data validation
## set data validation referencing range on separate sheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

sheet1->GetRange(L"B10")->SetText(L"Here is a dataValidation example.");

//This is the second sheet
intrusive_ptr<Worksheet> sheet2 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

//The property is to enable the data can be from different sheet.
sheet2->GetParentWorkbook()->SetAllow3DRangesInDataValidation(true);
sheet1->GetRange(L"B11")->GetDataValidation()->SetDataRange(dynamic_pointer_cast<CellRange>(sheet2->GetRange(L"A1:A7")));
```

---

# spire.xls cpp time validation
## create time data validation in Excel cell
```cpp
// Set text for cell C12
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12"))->SetText(L"Please enter time between 09:00 and 18:00:");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12"))->AutoFitColumns();

// Set Time data validation for cell D12
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D12"));
range->GetDataValidation()->SetAllowType(CellDataType::Time);
range->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);

range->GetDataValidation()->SetFormula1(L"09:00");
range->GetDataValidation()->SetFormula2(L"18:00");

range->GetDataValidation()->SetAlertStyle(AlertStyleType::Info);
range->GetDataValidation()->SetShowError(true);
range->GetDataValidation()->SetErrorTitle(L"Time Error");
range->GetDataValidation()->SetErrorMessage(L"Please enter a valid time");
range->GetDataValidation()->SetInputMessage(L"Time Validation Type");
range->GetDataValidation()->SetIgnoreBlank(true);
range->GetDataValidation()->SetShowInput(true);
```

---

# c++ excel data validation
## verify cell data against validation rules
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Cell B4 has the Decimal Validation
intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"));

//Get the validation of this cell
intrusive_ptr<Validation> validation = cell->GetDataValidation();

//Get the specified data range
double minimum = std::stod(validation->GetFormula1());
double maximum = std::stod(validation->GetFormula2());

//Set different numbers for the cell
for (int i = 5; i < 100; i = i + 40)
{
    cell->SetNumberValue(i);
    std::wstring result = L"";
    //Verify 
    if (cell->GetNumberValue() < minimum || cell->GetNumberValue() > maximum)
    {
        //Set string format for displaying
        result = L"Is input " + std::to_wstring(i) + L" a valid value for this Cell: false";
    }
    else
    {
        //Set string format for displaying
        result = L"Is input " + std::to_wstring(i) + L" a valid value for this Cell: true";
    }
    //Add result string to StringBuilder
    content->append(result);
}
```

---

# spire.xls cpp data validation
## implement whole number data validation for Excel cells
```cpp
//Set text for cell C12
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12"))->SetText(L"Please enter number between 10 and 100:");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12"))->AutoFitColumns();

//Set Whole Number data validation for cell D12
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D12"));
range->GetDataValidation()->SetAllowType(CellDataType::Integer);
range->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);

range->GetDataValidation()->SetFormula1(L"10");
range->GetDataValidation()->SetFormula2(L"100");

range->GetDataValidation()->SetAlertStyle(AlertStyleType::Info);
range->GetDataValidation()->SetShowError(true);
range->GetDataValidation()->SetErrorTitle(L"Error");
range->GetDataValidation()->SetErrorMessage(L"Please enter a valid number");
range->GetDataValidation()->SetInputMessage(L"Whole Number Validation Type");
range->GetDataValidation()->SetIgnoreBlank(true);
range->GetDataValidation()->SetShowInput(true);
```

---

# spire.xls cpp chart
## add data table to chart
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first chart
intrusive_ptr<Spire::Xls::Chart> chart = dynamic_pointer_cast<Spire::Xls::Chart>(dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0)));
chart->SetHasDataTable(true);
```

---

# spire.xls cpp chart error bars
## Add error bars to charts in Excel
```cpp
//Add a line chart and then add percentage error bar to the chart
intrusive_ptr<Spire::Xls::Chart> chart = sheet->GetCharts()->Add(ExcelChartType::Line);
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1:B7")));
chart->SetSeriesDataFromRange(false);
//Set chart position
chart->SetTopRow(8);
chart->SetBottomRow(25);
chart->SetLeftColumn(2);
chart->SetRightColumn(9);
chart->SetChartTitle(L"Error Bar 10% Plus");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);
intrusive_ptr<ChartSerie> cs1 = chart->GetSeries()->Get(0);
cs1->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A7")));
cs1->ErrorBar(true, ErrorBarIncludeType::Plus, ErrorBarType::Percentage, 10);

// Add a column chart with standard error bars as comparison
intrusive_ptr<Spire::Xls::Chart> chart2 = dynamic_pointer_cast<Spire::Xls::Chart>(sheet->GetCharts()->Add(ExcelChartType::ColumnClustered));
chart2->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1:C7")));
chart2->SetSeriesDataFromRange(false);
//Set chart position
chart2->SetTopRow(8);
chart2->SetBottomRow(25);
chart2->SetLeftColumn(10);
chart2->SetRightColumn(17);
chart2->SetChartTitle(L"Standard Error Bar");
chart2->GetChartTitleArea()->SetIsBold(true);
chart2->GetChartTitleArea()->SetSize(12);
intrusive_ptr<ChartSerie> cs2 = chart2->GetSeries()->Get(0);
cs2->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A7")));
cs2->ErrorBar(true, ErrorBarIncludeType::Minus, ErrorBarType::StandardError, 0.3);
intrusive_ptr<ChartSerie> cs3 = chart2->GetSeries()->Get(1);
cs3->ErrorBar(true, ErrorBarIncludeType::Both, ErrorBarType::StandardError, 0.5);
```

---

# spire.xls cpp chart
## add picture to chart
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Add the picture in chart
chart->GetShapes()->AddPicture(L"E-iceblueLogo.png");
```

---

# spire.xls cpp chart textbox
## add textbox to chart
```cpp
//Get the first chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Add a Textbox
intrusive_ptr<ITextBoxLinkShape> textbox = chart->GetShapes()->AddTextBox();
textbox->SetWidth(1200);
textbox->SetHeight(320);
textbox->SetLeft(1000);
textbox->SetTop(480);
textbox->SetText(L"This is a textbox");
```

---

# spire.xls cpp chart trendline
## add different types of trendlines to charts
```cpp
// Add logarithmic trendline to the first chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));
chart->SetChartTitle(L"Logarithmic Trendline");
chart->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Logarithmic);

// Add moving average trendline to the second chart
intrusive_ptr<Chart> chart1 = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(1));
chart1->SetChartTitle(L"Moving Average Trendline");
chart1->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Moving_Average);

// Add linear trendline to the third chart
intrusive_ptr<Chart> chart2 = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(2));
chart2->SetChartTitle(L"Linear Trendline");
chart2->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Linear);

// Add exponential trendline to the fourth chart
intrusive_ptr<Chart> chart3 = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(3));
chart3->SetChartTitle(L"Exponential Trendline");
chart3->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Exponential);
```

---

# spire.xls cpp chart adjust bar space
## adjust the gap width and overlap between bars in a chart
```cpp
//Get the first worksheet from workbook and then get the first chart from the worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Adjust the space between bars
for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
	intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
	cs->GetFormat()->GetOptions()->SetGapWidth(200);
	cs->GetFormat()->GetOptions()->SetOverlap(0);
}
```

---

# spire.xls cpp chart effect
## apply soft edges effect to chart
```cpp
//Get the chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Specify the size of the soft edge. Value can be set from 0 to 100
dynamic_pointer_cast<ChartArea>(chart->GetChartArea())->GetShadow()->SetSoftEdge(25);
```

---

# c++ chart size and position
## change chart size and position in excel
```cpp
//Get the chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Change chart size
chart->SetWidth(600);
chart->SetHeight(500);

//Change chart position
chart->SetLeftColumn(3);
chart->SetTopRow(7);
```

---

# spire.xls cpp chart
## change data label in chart
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the chart
intrusive_ptr<Spire::Xls::Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Change data label of the first datapoint of the first series
chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels()->SetText(L"changed data label");
```

---

# spire.xls cpp chart
## change chart data range
```cpp
//Get chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Change data range
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C4")));
```

---

# spire.xls cpp chart gridlines
## change major gridlines color
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Change the color of major gridlines
chart->GetPrimaryValueAxis()->GetMajorGridLines()->GetLineProperties()->SetColor(Spire::Xls::Color::GetRed());
```

---

# spire.xls cpp chart series color
## change chart series color
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Get the second series
intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(1);

//Set the fill type
cs->GetFormat()->GetFill()->SetFillType(ShapeFillType::SolidColor);

//Change the fill color
cs->GetFormat()->GetFill()->SetForeColor(Spire::Common::Color::GetOrange());
```

---

# Chart Axis Title Configuration
## Set titles and font size for chart axes
```cpp
//Set axis title
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Category Axis");
chart->GetPrimaryValueAxis()->SetTitle(L"Value axis");

//Set font size
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetSize(12);
chart->GetPrimaryValueAxis()->GetFont()->SetSize(12);
```

---

# c++ chart to image conversion
## convert excel chart to image file
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<Stream> image = workbook->SaveChartAsImage(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0)), 0);
image->Save(outputFile.c_str());
image->Close();

//Dispose resources
workbook->Dispose();
```

---

# spire.xls cpp clustered bar chart
## create a clustered bar chart with customized axes and legend
```cpp
//Add a chart
intrusive_ptr<Spire::Xls::Chart> chart = sheet->GetCharts()->Add();

//Set region of chart data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5")));
chart->SetSeriesDataFromRange(false);

//Set position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(6);
chart->SetRightColumn(11);
chart->SetBottomRow(29);
chart->SetChartType(ExcelChartType::BarClustered);

//Chart title
chart->SetChartTitle(L"Sales market by country");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Country");
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetIsBold(true);
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetTitleArea()->SetIsBold(true);
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetTitleArea()->SetTextRotationAngle(90);

chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
chart->GetPrimaryValueAxis()->SetMinValue(1000);
chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);

for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
    intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
    cs->GetFormat()->GetOptions()->SetIsVaryColor(true);
    cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
}

chart->GetLegend()->SetPosition(LegendPositionType::Top);
```

---

# spire.xls cpp chart
## create Clustered Column chart
```cpp
//Add a chart to the sheet
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();

//Set data range of chart 
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5")));
chart->SetSeriesDataFromRange(false);

//Set position of the chart
chart->SetLeftColumn(1);
chart->SetTopRow(6);
chart->SetRightColumn(11);
chart->SetBottomRow(29);

chart->SetChartType(ExcelChartType::ColumnClustered);

//Chart title
chart->SetChartTitle(L"Sales market by country");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

//Chart Axis
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Country");
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetIsBold(true);
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetTitleArea()->SetIsBold(true);

chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
chart->GetPrimaryValueAxis()->SetMinValue(1000);
chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);
chart->GetPrimaryValueAxis()->GetTitleArea()->SetTextRotationAngle(90);

for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
    intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
    cs->GetFormat()->GetOptions()->SetIsVaryColor(true);
    cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
}

//Chart Legend
chart->GetLegend()->SetPosition(LegendPositionType::Top);
```

---

# spire.xls cpp chart
## create Box and Whisker chart
```cpp
// Add a new chart to the worksheet.
auto officeChart = sheet->GetCharts()->Add();

// Set the title of the chart.
officeChart->SetChartTitle(L"Yearly Vehicle Sales");

// Set the chart type to Box and Whisker.
officeChart->SetChartType(ExcelChartType::BoxAndWhisker);

// Set the data range for the chart to range A1:E17.
officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E17")));

// Get the first series from the chart's series collection.
auto seriesA = officeChart->GetSeries()->Get(0);
seriesA->GetDataFormat()->SetShowInnerPoints(false);
seriesA->GetDataFormat()->SetShowOutlierPoints(true);
seriesA->GetDataFormat()->SetShowMeanMarkers(true);
seriesA->GetDataFormat()->SetShowMeanLine(false);
seriesA->GetDataFormat()->SetQuartileCalculationType(ExcelQuartileCalculation::ExclusiveMedian);

// Get the second series from the chart's series collection.
auto seriesB = officeChart->GetSeries()->Get(1);
seriesB->GetDataFormat()->SetShowInnerPoints(false);
seriesB->GetDataFormat()->SetShowOutlierPoints(true);
seriesB->GetDataFormat()->SetShowMeanMarkers(true);
seriesB->GetDataFormat()->SetShowMeanLine(false);
seriesB->GetDataFormat()->SetQuartileCalculationType(ExcelQuartileCalculation::InclusiveMedian);

// Get the third series from the chart's series collection.
auto seriesC = officeChart->GetSeries()->Get(2);
seriesC->GetDataFormat()->SetShowInnerPoints(false);
seriesC->GetDataFormat()->SetShowOutlierPoints(true);
seriesC->GetDataFormat()->SetShowMeanMarkers(true);
seriesC->GetDataFormat()->SetShowMeanLine(false);
seriesC->GetDataFormat()->SetQuartileCalculationType(ExcelQuartileCalculation::ExclusiveMedian);
```

---

# spire.xls cpp bubble chart
## create bubble chart in Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add a chart
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::Bubble);

//Set region of chart data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5")));
chart->SetSeriesDataFromRange(false);
chart->GetSeries()->Get(0)->SetBubbles(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:C5")));

//Set position of chart
chart->SetLeftColumn(7);
chart->SetTopRow(6);
chart->SetRightColumn(16);
chart->SetBottomRow(29);

chart->SetChartTitle(L"Bubble Chart");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);
```

---

# spire.xls cpp pivot chart
## create chart based on pivot table
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));

dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetCharts()->Add(ExcelChartType::BarClustered, pt);
```

---

# spire.xls cpp custom chart
## create a custom chart with different chart types for different series
```cpp
//Add a chart based on the data from A1 to B4
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B4")));
chart->SetSeriesDataFromRange(false);

//Set position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(10);
chart->SetRightColumn(7);
chart->SetBottomRow(25);

//Apply different chart type to different series
auto cs1 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Get(0));
cs1->SetSerieType(ExcelChartType::ColumnClustered);
auto cs2 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Get(1));
cs2->SetSerieType(ExcelChartType::Line);

chart->SetChartTitle(L"Custom chart");
```

---

# spire.xls cpp chart
## create Doughnut chart
```cpp
//Add a new chart, set chart type as doughnut
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();
chart->SetChartType(ExcelChartType::Doughnut);
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B5")));
chart->SetSeriesDataFromRange(false);

//Set position of chart
chart->SetLeftColumn(4);
chart->SetTopRow(2);
chart->SetRightColumn(12);
chart->SetBottomRow(22);

//Chart title
chart->SetChartTitle(L"Market share by country");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
    intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
    cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasPercentage(true);
}

chart->GetLegend()->SetPosition(LegendPositionType::Top);
```

---

# spire.xls cpp chart
## create funnel chart
```cpp
using namespace Spire::Xls;

// Create a new Workbook object
intrusive_ptr<Workbook> workbook = new Workbook();

// Retrieve the first worksheet from the workbook
auto sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Add a new chart to the worksheet
auto officeChart = sheet->GetCharts()->Add();

// Set the chart type to Funnel
officeChart->SetChartType(ExcelChartType::Funnel);

// Set the data range for the chart to range A1:B6
officeChart->SetDataRange(sheet->GetRange(L"A1:B6"));

// Set the title of the chart
officeChart->SetChartTitle(L"Funnel");

// Disable the legend in the chart
officeChart->SetHasLegend(false);

// Enable data labels for the default data point of the first series
officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);

// Set the size of the data labels to 8 points
officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(8);
```

---

# spire.xls cpp histogram chart
## create and configure a histogram chart in excel using c++
```cpp
// Add a new chart to the worksheet and cast it to Chart type.
auto officeChart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Add());

// Set the chart type to Histogram.
officeChart->SetChartType(ExcelChartType::Histogram);

// Set the data range for the chart to column A (A1:A15).
officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:A15")));

// Set the top row, bottom row, left column, and right column of the chart's display area.
officeChart->SetTopRow(1);
officeChart->SetBottomRow(19);
officeChart->SetLeftColumn(4);
officeChart->SetRightColumn(12);

// Set the bin width for the chart's primary category axis (X-axis).
dynamic_pointer_cast<ChartCategoryAxis>(officeChart->GetPrimaryCategoryAxis())->SetBinWidth(8);

// Set the gap width between bars in the chart's series.
officeChart->GetSeries()->Get(0)->GetDataFormat()->GetOptions()->SetGapWidth(6);

// Set the title of the chart.
officeChart->SetChartTitle(L"Height Data");

// Set the title of the primary value axis (Y-axis).
officeChart->GetPrimaryValueAxis()->SetTitle(L"Number of students");

// Set the title of the primary category axis (X-axis).
officeChart->GetPrimaryCategoryAxis()->SetTitle(L"Height");

// Disable the legend in the chart.
officeChart->SetHasLegend(false);
```

---

# spire.xls cpp multilevel chart
## create multi-level category chart
```cpp
//Add a clustered bar chart to worksheet
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::BarClustered);
chart->SetChartTitle(L"Value");
dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetFillType(ShapeFillType::NoFill);
chart->GetLegend()->Delete();
chart->SetLeftColumn(5);
chart->SetTopRow(1);
chart->SetRightColumn(14);

//Set the data source of series data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:C9")));
chart->SetSeriesDataFromRange(false);
//Set the data source of category labels
intrusive_ptr<ChartSerie> serie = chart->GetSeries()->Get(0);
serie->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:B9")));
//Show multi-level category labels
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetMultiLevelLable(true);
```

---

# spire.xls cpp chart
## create Pareto chart
```cpp
// Add a new chart to the worksheet.
auto officeChart = sheet->GetCharts()->Add();

// Set the chart type to Pareto.
officeChart->SetChartType(ExcelChartType::Pareto);

// Set the data range for the chart to range A2:B8.
officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:B8")));

// Set the top row, bottom row, left column, and right column of the chart's display area.
officeChart->SetTopRow(1);
officeChart->SetBottomRow(19);
officeChart->SetLeftColumn(4);
officeChart->SetRightColumn(12);

// Get the primary category axis (X-axis) of the chart and cast it to ChartCategoryAxis.
auto axis = dynamic_pointer_cast<ChartCategoryAxis>(officeChart->GetPrimaryCategoryAxis());

// Enable binning by category on the primary category axis.
axis->SetIsBinningByCategory(true);

// Set the value to be used for overflow bins on the primary category axis.
axis->SetOverflowBinValue(5);

// Set the value to be used for underflow bins on the primary category axis.
axis->SetUnderflowBinValue(1);

// Set the line color of the Pareto line in the chart.
officeChart->GetSeries()->Get(0)->GetParetoLineFormat()->GetLineProperties()->SetColor(Color::GetBlue());

// Set the gap width between bars in the chart's series.
officeChart->GetSeries()->Get(0)->GetDataFormat()->GetOptions()->SetGapWidth(6);

// Set the title of the chart.
officeChart->SetChartTitle(L"Expenses");

// Disable the legend in the chart.
officeChart->SetHasLegend(false);
```

---

# spire.xls cpp pivot chart
## create pivot chart from pivot table
```cpp
//get the first pivot table in the worksheet
intrusive_ptr<IPivotTable> pivotTable = sheet->GetPivotTables()->Get(0);

//create a clustered column chart based on the pivot table
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::ColumnClustered, pivotTable);
//set chart position
chart->SetTopRow(10);
chart->SetLeftColumn(1);
chart->SetRightColumn(7);
chart->SetBottomRow(25);
//set chart title
chart->SetChartTitle(L"Pivot Chart");
```

---

# spire.xls cpp chart
## create radar chart
```cpp
//Add a new chart worsheet to workbook
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();

//Set position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(6);
chart->SetRightColumn(11);
chart->SetBottomRow(29);

//Set region of chart data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5")));
chart->SetSeriesDataFromRange(false);

chart->SetChartType(ExcelChartType::Radar);

//Chart title
chart->SetChartTitle(L"Sale market by region");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetVisible(false);

chart->GetLegend()->SetPosition(LegendPositionType::Corner);
```

---

# spire.xls cpp chart
## create SunBurst chart
```cpp
// Add a new chart to the worksheet.
auto officeChart = sheet->GetCharts()->Add();

// Set the chart type to Sunburst.
officeChart->SetChartType(ExcelChartType::SunBurst);

// Set the data range for the chart to range A1:D16.
officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D16")));

// Set the top row, bottom row, left column, and right column of the chart's display area.
officeChart->SetTopRow(1);
officeChart->SetBottomRow(17);
officeChart->SetLeftColumn(6);
officeChart->SetRightColumn(14);

// Set the title of the chart.
officeChart->SetChartTitle(L"Sales by quarter");

// Set the size of the data labels for the default data point of the first series to 8 points.
officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(8);

// Disable the legend in the chart.
officeChart->SetHasLegend(false);
```

---

# spire.xls cpp chart
## create TreeMap chart
```cpp
// Add a new chart to the worksheet.
auto officeChart = sheet->GetCharts()->Add();

// Set the chart type to TreeMap.
officeChart->SetChartType(ExcelChartType::TreeMap);

// Set the data range for the chart to range A2:C11.
officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:C11")));

// Set the top row, bottom row, left column, and right column of the chart's display area.
officeChart->SetTopRow(1);
officeChart->SetBottomRow(19);
officeChart->SetLeftColumn(4);
officeChart->SetRightColumn(14);

// Set the title of the chart.
officeChart->SetChartTitle(L"Area by countries");

// Set the TreeMap label option for the first series to Banner.
officeChart->GetSeries()->Get(0)->GetDataFormat()->SetTreeMapLabelOption(ExcelTreeMapLabelOption::Banner);

// Set the size of the data labels for the default data point of the first series to 8 points.
officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(8);
```

---

# spire.xls cpp chart
## create Waterfall chart
```cpp
// Create an intrusive pointer to a Workbook object
intrusive_ptr<Workbook>workbook = new Workbook();

// Get the first worksheet from the workbook
auto sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Add a chart to the worksheet
auto officeChart = sheet->GetCharts()->Add();

// Set the chart type to Waterfall
officeChart->SetChartType(ExcelChartType::WaterFall);

// Set the data range for the chart
officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:B8")));

// Set the top row, bottom row, left column, and right column for the chart
officeChart->SetTopRow(1);
officeChart->SetBottomRow(19);
officeChart->SetLeftColumn(4);
officeChart->SetRightColumn(12);

// Set certain data points in the chart as total
dynamic_pointer_cast<XlsChartDataPoint>(officeChart->GetSeries()->Get(0)->GetDataPoints()->Get(3))->SetSetAsTotal(true);
dynamic_pointer_cast<XlsChartDataPoint>(officeChart->GetSeries()->Get(0)->GetDataPoints()->Get(6))->SetSetAsTotal(true);

// Show connector lines in the chart
dynamic_pointer_cast<ChartSerieDataFormat>(officeChart->GetSeries()->Get(0)->GetFormat())->SetShowConnectorLines(true);

// Set the chart title
officeChart->SetChartTitle(L"WaterFall Chart");

// Enable data labels for the default data point of the chart series
officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(8);

// Set the position of the legend to the right
officeChart->GetLegend()->SetPosition(LegendPositionType::Right);
```

---

# spire.xls cpp chart
## customize data markers in scatter chart
```cpp
//Create a Scatter-Markers chart based on the sample data
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::ScatterMarkers);
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B7")));
dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->SetVisible(false);
chart->SetSeriesDataFromRange(false);
chart->SetTopRow(5);
chart->SetBottomRow(22);
chart->SetLeftColumn(4);
chart->SetRightColumn(11);
chart->SetChartTitle(L"Chart with Markers");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(10);

//Format the markers in the chart by setting the background color, foreground color, type, size and transparency
intrusive_ptr<ChartSerie> cs1 = chart->GetSeries()->Get(0);
cs1->GetDataFormat()->SetMarkerBackgroundColor(Spire::Xls::Color::GetRoyalBlue());
cs1->GetDataFormat()->SetMarkerForegroundColor(Spire::Xls::Color::GetWhiteSmoke());
cs1->GetDataFormat()->SetMarkerSize(7);
cs1->GetDataFormat()->SetMarkerStyle(ChartMarkerType::PlusSign);
cs1->GetDataFormat()->SetMarkerTransparencyValue(0.8);

intrusive_ptr<ChartSerie> cs2 = chart->GetSeries()->Get(1);
cs2->GetDataFormat()->SetMarkerBackgroundColor(Spire::Xls::Color::GetPink());
cs2->GetDataFormat()->SetMarkerSize(9);
cs2->GetDataFormat()->SetMarkerStyle(ChartMarkerType::Triangle);
cs2->GetDataFormat()->SetMarkerTransparencyValue(0.9);
```

---

# spire.xls cpp chart data callout
## configure chart data labels with callout in Excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
	intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
	cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
	(dynamic_pointer_cast<XlsChartDataLabels>(cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()))->SetHasWedgeCallout(true);
	cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasCategoryName(true);
	cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasSeriesName(true);
	cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasLegendKey(true);
}
```

---

# spire.xls cpp chart legend
## delete legend entries from chart
```cpp
//Get the chart
intrusive_ptr<Spire::Xls::Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Delete the first and the second legend entries from the chart
chart->GetLegend()->GetLegendEntries()->Get(0)->Delete();
chart->GetLegend()->GetLegendEntries()->Get(1)->Delete();
```

---

# spire.xls cpp chart with discontinuous data
## create a chart with non-continuous data ranges
```cpp
//Add a chart
intrusive_ptr<Spire::Xls::Chart> chart = sheet->GetCharts()->Add(ExcelChartType::ColumnClustered);
chart->SetSeriesDataFromRange(false);

//Set the position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(10);
chart->SetRightColumn(10);
chart->SetBottomRow(24);

//Add a series
intrusive_ptr<ChartSerie> cs1 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Add());

//Set the name of the cs1
cs1->SetName(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetValue());

//Set discontinuous values for cs1
cs1->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A3"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:A6"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:A9")))));
cs1->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B3"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5:B6"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8:B9")))));

//Set the chart type
cs1->SetSerieType(ExcelChartType::ColumnClustered);

//Add a series
intrusive_ptr<ChartSerie> cs2 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Add());
cs2->SetName(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->GetValue());
cs2->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A3"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:A6"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:A9")))));
cs2->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:C3"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5:C6"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C8:C9")))));
cs2->SetSerieType(ExcelChartType::ColumnClustered);

chart->SetChartTitle(L"Chart");
chart->GetChartTitleArea()->GetFont()->SetSize(20);
chart->GetChartTitleArea()->SetColor(Spire::Xls::Color::GetBlack());

chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
```

---

# spire.xls cpp chart
## edit line chart by adding new series
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the line chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Add a new series
intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Add(L"Added");

//Set the values for the series
cs->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"I1:L1")));
```

---

# spire.xls cpp chart
## create exploded doughnut chart
```cpp
//Add a chart
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();
chart->SetChartType(ExcelChartType::DoughnutExploded);

//Set position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(6);
chart->SetRightColumn(11);
chart->SetBottomRow(29);

//Set region of chart data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B5")));
chart->SetSeriesDataFromRange(false);

//Chart title
chart->SetChartTitle(L"Sales market by country");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
    intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
    cs->GetFormat()->GetOptions()->SetIsVaryColor(true);
    cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
}

dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetVisible(false);
chart->GetLegend()->SetPosition(LegendPositionType::Top);
```

---

# spire.xls cpp trendline
## extract trendline equation from chart
```cpp
//Get the chart from the first worksheet
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetCharts()->Get(0));

//Get the trendline of the chart and then extract the equation of the trendline
intrusive_ptr<IChartTrendLine> trendLine = chart->GetSeries()->Get(1)->GetTrendLines()->GetItem(0);
wstring formula = trendLine->GetFormula();
```

---

# spire.xls cpp chart
## fill chart elements with picture
```cpp
//Get the first chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));
//Fill chart area with image
chart->GetChartArea()->GetFill()->CustomPicture(new Stream(L"background.png"), L"None");

chart->GetPlotArea()->GetFill()->SetTransparency(0.9);
```

---

# spire.xls cpp chart axis formatting
## format chart axis properties and appearance
```cpp
//Add a chart
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::ColumnClustered);
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1:B9")));
chart->SetSeriesDataFromRange(false);
dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->SetVisible(false);
chart->SetTopRow(10);
chart->SetBottomRow(28);
chart->SetLeftColumn(2);
chart->SetRightColumn(10);
chart->SetChartTitle(L"Chart with Customized Axis");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);
intrusive_ptr<ChartSerie> cs1 = chart->GetSeries()->Get(0);
cs1->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A9")));

//Format axis
chart->GetPrimaryValueAxis()->SetMajorUnit(8);
chart->GetPrimaryValueAxis()->SetMinorUnit(2);
chart->GetPrimaryValueAxis()->SetMaxValue(50);
chart->GetPrimaryValueAxis()->SetMinValue(0);
chart->GetPrimaryValueAxis()->SetIsReverseOrder(false);
chart->GetPrimaryValueAxis()->SetMajorTickMark(TickMarkType::TickMarkOutside);
chart->GetPrimaryValueAxis()->SetMinorTickMark(TickMarkType::TickMarkInside);
chart->GetPrimaryValueAxis()->SetTickLabelPosition(TickLabelPositionType::TickLabelPositionNextToAxis);
chart->GetPrimaryValueAxis()->SetCrossesAt(0);

//Set NumberFormat
chart->GetPrimaryValueAxis()->SetNumberFormat(L"$#,##0");
chart->GetPrimaryValueAxis()->SetIsSourceLinked(false);

intrusive_ptr<ChartSerie> serie = chart->GetSeries()->Get(0);
intrusive_ptr<Spire::Xls::IEnumerator<XlsChartDataPoint>> ie = dynamic_pointer_cast<ChartDataPointsCollection>(serie->GetDataPoints())->GetEnumerator();
while (ie->MoveNext())
{
    intrusive_ptr<IChartDataPoint> dataPoint = ie->GetCurrent();
    //Format Series
    dataPoint->GetDataFormat()->GetFill()->SetFillType(ShapeFillType::SolidColor);
    dataPoint->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetLightGreen());

    //Set transparency
    dataPoint->GetDataFormat()->GetFill()->SetTransparency(0.3);
}
```

---

# spire.xls cpp gauge chart
## create a gauge chart using doughnut and pie charts
```cpp
//Add a Doughnut chart
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::Doughnut);
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:A5")));
chart->SetSeriesDataFromRange(false);
chart->SetHasLegend(true);

//Set the position of chart
chart->SetLeftColumn(2);
chart->SetTopRow(7);
chart->SetRightColumn(9);
chart->SetBottomRow(25);

//Get the series 1
intrusive_ptr<ChartSerie> cs1 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Get(L"Value"));
cs1->GetFormat()->GetOptions()->SetDoughnutHoleSize(60);
cs1->GetDataFormat()->GetOptions()->SetFirstSliceAngle(270);

//Set the fill color
cs1->GetDataPoints()->Get(0)->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetYellow());
cs1->GetDataPoints()->Get(1)->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetPaleVioletRed());
cs1->GetDataPoints()->Get(2)->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetDarkViolet());
cs1->GetDataPoints()->Get(3)->GetDataFormat()->GetFill()->SetVisible(false);

//Add a series with pie chart
intrusive_ptr<ChartSerie> cs2 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Add(L"Pointer", ExcelChartType::Pie));

//Set the value
cs2->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2:D4")));
cs2->SetUsePrimaryAxis(false);
cs2->GetDataPoints()->Get(0)->GetDataLabels()->SetHasValue(true);
cs2->GetDataFormat()->GetOptions()->SetFirstSliceAngle(270);
cs2->GetDataPoints()->Get(0)->GetDataFormat()->GetFill()->SetVisible(false);
cs2->GetDataPoints()->Get(1)->GetDataFormat()->GetFill()->SetFillType(ShapeFillType::SolidColor);
cs2->GetDataPoints()->Get(1)->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetBlack());
cs2->GetDataPoints()->Get(2)->GetDataFormat()->GetFill()->SetVisible(false);
```

---

# spire.xls cpp chart
## get category labels from chart
```cpp
//Get the cell range of the category labels
std::wstring* content = new std::wstring();
intrusive_ptr<IXLSRange> cr = dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetCategoryLabels();
for (int i = 0; i < cr->GetCells()->GetCount(); i++)
{
    auto cell = cr->GetCells()->GetItem(i);
    wstring value = cell->GetValue();
    content->append(value + L"\r\n");
}
```

---

# c++ get chart data point values
## Extract data point values from a chart in an Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Get the first series of the chart
intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(0);
for (int i = 0; i < cs->GetValues()->GetCells()->GetCount(); i++)
{
    wstring address = cs->GetValues()->GetCells()->GetItem(i)->GetRangeAddress();
    
    //Get the data point value
    wstring value = cs->GetValues()->GetCells()->GetItem(i)->GetValue();
}
```

---

# spire.xls cpp chart
## get worksheet of chart
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Access the first chart inside this worksheet
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Get its worksheet
intrusive_ptr<Worksheet> wSheet = chart->GetWorksheet();
```

---

# spire.xls cpp chart
## hide major gridlines in chart
```cpp
//Get the chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Hide major gridlines
chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
```

---

# spire.xls cpp chart
## create Line chart
```cpp
//Add a chart
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();
chart->SetChartType(ExcelChartType::Line);

//Set region of chart data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E5")));

//Set position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(6);
chart->SetRightColumn(11);
chart->SetBottomRow(29);

//Set chart title
chart->SetChartTitle(L"Sales market by country");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Month");
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetIsBold(true);
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetTitleArea()->SetIsBold(true);

chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
chart->GetPrimaryValueAxis()->GetTitleArea()->SetTextRotationAngle(90);
chart->GetPrimaryValueAxis()->SetMinValue(1000);
chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);

for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
    intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
    cs->GetFormat()->GetOptions()->SetIsVaryColor(true);
    cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
}

dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetVisible(false);

chart->GetLegend()->SetPosition(LegendPositionType::Top);
```

---

# Spire.XLS C++ Chart Droplines
## Enable droplines for a line chart in Excel
```cpp
// Create an intrusive pointer to a Workbook object
intrusive_ptr<Workbook>workbook = new Workbook();

// Get the first worksheet from the workbook
intrusive_ptr<Worksheet>worksheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Get the first chart from the worksheet
intrusive_ptr<Chart>chart = dynamic_pointer_cast<Chart>(worksheet->GetCharts()->Get(0));

// Enable droplines for the first series of the chart
chart->GetSeries()->Get(0)->SetHasDroplines(true);

// Dispose of the workbook object
workbook->Dispose();
```

---

# spire.xls cpp pie chart
## create a pie chart with data labels
```cpp
//Add a chart
intrusive_ptr<Chart> chart = nullptr;
chart = sheet->GetCharts()->Add(ExcelChartType::Pie);

//Set region of chart data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B5")));
chart->SetSeriesDataFromRange(false);

//Set position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(6);
chart->SetRightColumn(9);
chart->SetBottomRow(25);

//Chart title
chart->SetChartTitle(L"Sales by year");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(0);
cs->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A5")));
cs->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B5")));
cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);

dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetVisible(false);
```

---

# spire.xls cpp pyramid column chart
## Create a 3D clustered pyramid column chart with customized axes and legend
```cpp
//Add a chart
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();

//Set region of chart data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B5")));
chart->SetSeriesDataFromRange(false);

//Set position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(6);
chart->SetRightColumn(11);
chart->SetBottomRow(29);

chart->SetChartType(ExcelChartType::Pyramid3DClustered);

//Chart title
chart->SetChartTitle(L"Sales by year");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Year");
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetIsBold(true);
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetTitleArea()->SetIsBold(true);

chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
chart->GetPrimaryValueAxis()->SetMinValue(1000);
chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);
chart->GetPrimaryValueAxis()->GetTitleArea()->SetTextRotationAngle(90);

intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(0);
cs->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A5")));
cs->GetFormat()->GetOptions()->SetIsVaryColor(true);

chart->GetLegend()->SetPosition(LegendPositionType::Top);
```

---

# spire.xls cpp chart removal
## Remove a chart from Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first chart from the first worksheet
intrusive_ptr<IChartShape> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));
//Remove the chart
chart->Remove();
```

---

# spire.xls cpp chart resize move
## Resize and move chart in Excel worksheet
```cpp
//Get the chart from the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Set position of the chart
chart->SetLeftColumn(5);
chart->SetTopRow(1);

//Resize the chart
chart->SetWidth(500);
chart->SetHeight(350);
```

---

# spire.xls cpp chart
## apply rich text formatting to chart data labels
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first chart inside this worksheet
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Get the first datalabel of the first series 
intrusive_ptr<IChartDataLabels> datalabel = chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels();

//Set the text
datalabel->SetText(L"Rich Text Label");

//Show the value
chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels()->SetHasValue(true);

//Set styles for the text
chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels()->SetColor(Spire::Xls::Color::GetRed());
chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels()->SetIsBold(true);
```

---

# spire.xls cpp 3d chart rotation
## Rotate a 3D chart by setting X rotation and Y rotation values
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//X rotation:
chart->SetRotation(30);
//Y rotation:
chart->SetElevation(20);
```

---

# spire.xls cpp scatter chart
## create scatter chart with trend line
```cpp
//Add a chart
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::ScatterMarkers);

//Set region of chart data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B7")));
chart->SetSeriesDataFromRange(false);

//Set position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(11);
chart->SetRightColumn(10);
chart->SetBottomRow(28);

chart->SetChartTitle(L"Scatter Chart");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

chart->GetSeries()->Get(0)->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A7")));
chart->GetSeries()->Get(0)->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B7")));

//Add a trend line for the first series
chart->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Exponential);

chart->GetPrimaryValueAxis()->SetTitle(L"Month");
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Planned");
```

---

# spire.xls cpp chart data labels
## Set and format data labels in Excel chart
```cpp
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::LineMarkers);
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1:B7")));
dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->SetVisible(false);
chart->SetSeriesDataFromRange(false);
chart->SetTopRow(5);
chart->SetBottomRow(26);
chart->SetLeftColumn(2);
chart->SetRightColumn(11);
chart->SetChartTitle(L"Data Labels Demo");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);
intrusive_ptr<ChartSerie> cs1 = chart->GetSeries()->Get(0);
cs1->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A7")));

cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasLegendKey(false);
cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasPercentage(false);
cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasSeriesName(true);
cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasCategoryName(true);
cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetDelimiter(L". ");

cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(9);
cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetColor(Spire::Xls::Color::GetRed());
cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetFontName(L"Calibri");
cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetPosition(DataLabelPositionType::Center);
```

---

# spire.xls cpp chart
## set border color and style for chart
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Set CustomLineWeight property for Series line
(dynamic_pointer_cast<XlsChartBorder>(chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataFormat()->GetLineProperties()))->SetCustomLineWeight(2.5f);
//Set color property for Series line
(dynamic_pointer_cast<XlsChartBorder>(chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataFormat()->GetLineProperties()))->SetColor(Spire::Xls::Color::GetRed());
```

---

# spire.xls cpp chart marker
## set border width of chart markers
```cpp
//Get the chart from the first worksheet
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetCharts()->Get(0));

chart->GetSeries()->Get(0)->GetDataFormat()->SetMarkerBorderWidth(1.5); //unit is pt

chart->GetSeries()->Get(1)->GetDataFormat()->SetMarkerBorderWidth(2.5); //unit is pt
```

---

# c++ chart background color
## Set chart background color in Excel using Spire.XLS
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Set background color
dynamic_pointer_cast<ChartArea>(chart->GetChartArea())->SetForeGroundColor(Spire::Xls::Color::GetLightYellow());
```

---

# spire.xls cpp chart color
## set color for chart area and plot area
```cpp
//Get the chart
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Set color for chart area
dynamic_pointer_cast<ChartArea>(chart->GetChartArea())->GetFill()->SetForeColor(Spire::Xls::Color::GetLightSeaGreen());

//Set color for plot area
dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetForeColor(Spire::Xls::Color::GetLightGray());
```

---

# spire.xls cpp font
## set font for chart data labels
```cpp
//Create a font
intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
font->SetSize(15.0);
font->SetColor(Spire::Xls::Color::GetLightSeaGreen());

for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
    intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
    //Set font
    (dynamic_pointer_cast<XlsChartDataLabels>(cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()))->GetTextArea()->SetFont(font);
}
```

---

# c++ chart font styling
## set font for chart legend and data table
```cpp
//Create a font with specified size and color
intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
font->SetSize(14.0);
font->SetColor(Spire::Xls::Color::GetRed());

//Apply the font to chart Legend
(dynamic_pointer_cast<ChartTextArea>(chart->GetLegend()->GetTextArea()))->SetFont(font);

//Apply the font to chart DataLabel
for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
    intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
    (dynamic_pointer_cast<XlsChartDataLabels>(cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()))->GetTextArea()->SetFont(font);
}
```

---

# spire.xls cpp chart font
## set font properties for chart title and axes
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

//Format the font for the chart title
chart->GetChartTitleArea()->SetColor(Spire::Xls::Color::GetBlue());
chart->GetChartTitleArea()->SetSize(20.0);

//Format the font for the chart Axis
chart->GetPrimaryValueAxis()->GetFont()->SetColor(Spire::Xls::Color::GetGold());
chart->GetPrimaryValueAxis()->GetFont()->SetSize(10.0);
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetColor(Spire::Xls::Color::GetRed());
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetSize(20.0);
```

---

# spire.xls cpp chart legend
## set legend background color
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

intrusive_ptr<XlsChartFrameFormat> x = dynamic_pointer_cast<XlsChartFrameFormat>(dynamic_pointer_cast<ChartLegend>(chart->GetLegend())->GetFrameFormat());
x->GetFill()->SetFillType(ShapeFillType::SolidColor);
x->SetForeGroundColor(Spire::Xls::Color::GetSkyBlue());
```

---

# spire.xls cpp trendline
## set number format of trendline in chart
```cpp
//Get the chart from the first worksheet
intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetCharts()->Get(0));

//Get the trendline of the chart and then extract the equation of the trendline
intrusive_ptr<IChartTrendLine> trendLine = chart->GetSeries()->Get(1)->GetTrendLines()->GetItem(0);

//Set the number format of trendLine to "#,##0.00"
trendLine->GetDataLabel()->SetNumberFormat(L"#,##0.00");
```

---

# spire.xls cpp chart
## show leader lines for chart data labels
```cpp
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::BarStacked);
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C3")));
chart->SetTopRow(4);
chart->SetLeftColumn(2);
chart->SetWidth(450);
chart->SetHeight(300);

for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
	intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
	cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
	cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetShowLeaderLines(true);
}
```

---

# spire.xls cpp sparkline
## add sparklines to excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add sparkline
intrusive_ptr<SparklineGroup> sparklineGroup = sheet->GetSparklineGroups()->AddGroup(SparklineType::Line);
intrusive_ptr<SparklineCollection> sparklines = sparklineGroup->Add();
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:D2")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E2")));
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3:D3")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E3")));
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4:D4")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E4")));
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:D5")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E5")));
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6:D6")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E6")));
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7:D7")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E7")));
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:D8")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E8")));
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A9:D9")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E9")));
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A10:D10")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E10")));
sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A11:D11")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E11")));
```

---

# spire.xls cpp stacked column chart
## create a stacked column chart with customized axes and legend
```cpp
//Add a chart
intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();

//Set region of chart data
chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5")));
chart->SetSeriesDataFromRange(false);

//Set position of chart
chart->SetLeftColumn(1);
chart->SetTopRow(6);
chart->SetRightColumn(11);
chart->SetBottomRow(29);
chart->SetChartType(ExcelChartType::ColumnStacked);

//Chart title
chart->SetChartTitle(L"Sales market by country");
chart->GetChartTitleArea()->SetIsBold(true);
chart->GetChartTitleArea()->SetSize(12);

//Chart Axes
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Country");
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetIsBold(true);
dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetTitleArea()->SetIsBold(true);

chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
chart->GetPrimaryValueAxis()->SetMinValue(1000);
chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);
chart->GetPrimaryValueAxis()->GetTitleArea()->SetTextRotationAngle(90);

for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
{
    intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
    cs->GetFormat()->GetOptions()->SetIsVaryColor(true);
    cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
}

//Chart Legend
chart->GetLegend()->SetPosition(LegendPositionType::Top);
```

---

# spire.xls cpp shapes
## add arrow lines to excel worksheet
```cpp
//Add a Double Arrow and fill the line with solid color.
auto line = sheet->GetTypedLines()->AddLine();
line->SetTop(10);
line->SetLeft(20);
line->SetWidth(100);
line->SetHeight(0);
line->SetColor(Spire::Xls::Color::GetBlue());
line->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrow);
line->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);
//Add an Arrow and fill the line with solid color.
auto line_1 = sheet->GetTypedLines()->AddLine();
line_1->SetTop(50);
line_1->SetLeft(30);
line_1->SetWidth(100);
line_1->SetHeight(100);
line_1->SetColor(Spire::Xls::Color::GetRed());
line_1->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineNoArrow);
line_1->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);

//Add an Elbow Arrow Connector.
intrusive_ptr<XlsLineShape> line3 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
line3->SetLineShapeType(LineShapeType::ElbowLine);
line3->SetWidth(30);
line3->SetHeight(50);
line3->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);
line3->SetTop(100);
line3->SetLeft(50);

//Add an Elbow Double-Arrow Connector.
intrusive_ptr<XlsLineShape> line2 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
line2->SetLineShapeType(LineShapeType::ElbowLine);
line2->SetWidth(50);
line2->SetHeight(50);
line2->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);
line2->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrow);
line2->SetLeft(120);
line2->SetTop(100);

//Add a Curved Arrow Connector.
line3 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
line3->SetLineShapeType(LineShapeType::CurveLine);
line3->SetWidth(30);
line3->SetHeight(50);
line3->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrowOpen);
line3->SetTop(100);
line3->SetLeft(200);

//Add a Curved Double-Arrow Connector.
line2 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
line2->SetLineShapeType(LineShapeType::CurveLine);
line2->SetWidth(30);
line2->SetHeight(50);
line2->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrowOpen);
line2->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrowOpen);
line2->SetLeft(250);
line2->SetTop(100);
```

---

# spire.xls cpp shapes
## Add line shapes to Excel worksheet
```cpp
//Add shape line1
intrusive_ptr<ILineShape> line1 = sheet->GetLines()->AddLine(10, 2, 200, 1, LineShapeType::Line);
//Set dash style type
line1->SetDashStyle(ShapeDashLineStyleType::Solid);
//Set color
line1->SetColor(Spire::Xls::Color::GetCadetBlue());
//Set weight
line1->SetWeight(2.0f);
//Set end arrow style type
line1->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);

//Add shape line2
intrusive_ptr<ILineShape> line2 = sheet->GetLines()->AddLine(12, 2, 200, 1, LineShapeType::CurveLine);
line2->SetDashStyle(ShapeDashLineStyleType::Dotted);
line2->SetColor(Spire::Xls::Color::GetOrangeRed());
line2->SetWeight(2.0f);

//Add shape line3
intrusive_ptr<ILineShape> line3 = sheet->GetLines()->AddLine(14, 2, 200, 1, LineShapeType::ElbowLine);
line3->SetDashStyle(ShapeDashLineStyleType::DashDotDot);
line3->SetColor(Spire::Xls::Color::GetPurple());
line3->SetWeight(2.0f);

//Add shape line4
intrusive_ptr<ILineShape> line4 = sheet->GetLines()->AddLine(16, 2, 200, 1, LineShapeType::LineInv);
line4->SetDashStyle(ShapeDashLineStyleType::Dashed);
line4->SetColor(Spire::Xls::Color::GetGreen());
line4->SetWeight(2.0f);
```

---

# spire.xls cpp shapes
## add oval shapes to Excel worksheet
```cpp
//Add oval shape1
intrusive_ptr<IOvalShape> ovalShape1 = sheet->GetOvalShapes()->AddOval(11, 2, 100, 100);
ovalShape1->GetLine()->SetWeight(0);
//Fill shape with solid color
ovalShape1->GetFill()->SetFillType(ShapeFillType::SolidColor);
ovalShape1->GetFill()->SetForeColor(Spire::Xls::Color::GetDarkCyan());

//Add oval shape2
intrusive_ptr<IOvalShape> ovalShape2 = sheet->GetOvalShapes()->AddOval(11, 5, 100, 100);
ovalShape2->GetLine()->SetWeight(1);
//Fill shape with picture
ovalShape2->GetLine()->SetDashStyle(ShapeDashLineStyleType::Solid);
ovalShape2->GetFill()->CustomPicture(L"Logo.png");
```

---

# spire.xls cpp shapes
## add rectangle shapes to worksheet
```cpp
//Add rectangle shape 1------Rect
intrusive_ptr<IRectangleShape> rect1 = sheet->GetRectangleShapes()->AddRectangle(11, 2, 60, 100, RectangleShapeType::Rect);
rect1->GetLine()->SetWeight(1);
//Fill shape with solid color
rect1->GetFill()->SetFillType(ShapeFillType::SolidColor);
rect1->GetFill()->SetForeColor(Spire::Xls::Color::GetDarkGreen());

//Add rectangle shape 2------RoundRect
intrusive_ptr<IRectangleShape> rect2 = sheet->GetRectangleShapes()->AddRectangle(11, 5, 60, 100, RectangleShapeType::RoundRect);
rect2->GetLine()->SetWeight(1);
rect2->GetFill()->SetFillType(ShapeFillType::SolidColor);
rect2->GetFill()->SetForeColor(Spire::Xls::Color::GetDarkCyan());
```

---

# spire xls cpp spinner control
## add spinner control to excel worksheet
```cpp
//Add spinner control
intrusive_ptr<ISpinnerShape> spinner = sheet->GetSpinnerShapes()->AddSpinner(12, 4, 20, 20);
spinner->SetLinkedCell(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12")));
spinner->SetMin(0);
spinner->SetMax(100);
spinner->SetIncrementalChange(5);
spinner->SetDisplay3DShading(true);
```

---

# spire.xls cpp shapes
## adjust arrow polyline position
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Draw an elbow arrow
intrusive_ptr<XlsLineShape> line = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine(5, 5, 100, 100, LineShapeType::ElbowLine));
line->SetEndArrowHeadStyle(ShapeArrowStyleType::LineNoArrow);
line->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrow);
intrusive_ptr<GeomertyAdjustValue> ad = line->GetShapeAdjustValues()->AddAdjustValue(GeomertyAdjustValueFormulaType::LiteralValue);

//When the parameter value is less than 0, the focus of the line is on the left side of the left point, when it is equal to 0, the position is the same as the left point, it is equal to 50 in the middle of the graph, and when it is equal to 100, it is the same as the right point.
ad->SetFormulaParameter(std::vector<double> {-50});
```

---

# spire.xls cpp shapes
## copy shapes between worksheets
```cpp
//Get worksheets
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
intrusive_ptr<Worksheet> CopyShapes = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

//Create and copy line shape
auto line = sheet->GetTypedLines()->AddLine();
line->SetTop(50);
line->SetLeft(30);
line->SetWidth(30);
line->SetHeight(50);
line->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrowDiamond);
line->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);
CopyShapes->GetTypedLines()->AddCopy(line);

//Create and copy button
auto button = sheet->GetTypedRadioButtons()->Add(5, 5, 20, 20);
CopyShapes->GetTypedRadioButtons()->AddCopy(button);

//Create and copy textbox
auto textbox = sheet->GetTypedTextBoxes()->AddTextBox(5, 7, 50, 100);
CopyShapes->GetTypedTextBoxes()->AddCopy(textbox);

//Create and copy checkbox
auto checkbox = sheet->GetTypedCheckBoxes()->AddCheckBox(10, 1, 20, 20);
CopyShapes->GetTypedCheckBoxes()->AddCopy(checkbox);

//Create and copy combobox
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A14"))->SetValue(L"1");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A15"))->SetValue(L"2");
auto ComboBoxes = sheet->GetTypedComboBoxes()->AddComboBox(10, 5, 30, 30);
ComboBoxes->SetListFillRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A14:A15")));
CopyShapes->GetTypedComboBoxes()->AddCopy(ComboBoxes);
```

---

# spire.xls cpp shapes
## delete all shapes in worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Delete all shapes in the worksheet
for (int i = sheet->GetPrstGeomShapes()->GetCount() - 1; i >= 0; i--)
{
    sheet->GetPrstGeomShapes()->Get(i)->Remove();
}
```

---

# C++ Delete Excel Shape
## Delete a particular shape from an Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Delete the first shape in the worksheet
sheet->GetPrstGeomShapes()->Get(0)->Remove();
```

---

# spire.xls cpp shapes
## draw lines through two points in Excel
```cpp
//Draw a line according to relative position
intrusive_ptr<XlsLineShape> line1 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
line1->SetLeftColumn(3);
line1->SetTopRow(3);
line1->SetLeftColumnOffset(0);
line1->SetTopRowOffset(0);

line1->SetRightColumn(4);
line1->SetBottomRow(5);
line1->SetRightColumnOffset(0);
line1->SetBottomRowOffset(0);

//Draw a line according to absolute position(pixels)
intrusive_ptr<XlsLineShape> line2 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
intrusive_ptr<Point> startPoint = new Point();
startPoint->SetX(30), startPoint->SetY(50);
line2->SetStartPoint(startPoint);
intrusive_ptr<Point> endPoint = new Point();
endPoint->SetX(20), endPoint->SetY(80);
line2->SetEndPoint(endPoint);
```

---

# spire.xls cpp shape text extraction
## extract text from shapes in excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Extract text from the first shape and save to a txt file.
intrusive_ptr<IPrstGeomShape> shape1 = sheet->GetPrstGeomShapes()->Get(2);
wstring s = shape1->GetText();
wstring* content = new wstring();
content->append(L"The text in the third shape is: " + s);
```

---

# spire.xls cpp shapes
## get cell range linked to shapes
```cpp
// Get the first worksheet from the workbook
intrusive_ptr<Worksheet>sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Get the collection of preset geometric shapes from the worksheet
intrusive_ptr<PrstGeomShapeCollection>prstGeomShapeCollection = sheet->GetPrstGeomShapes();

// Get the shape with the name "Yesterday" from the collection
auto shape = prstGeomShapeCollection->Get(L"Yesterday");

// Get the cell address linked to the shape
std::wstring cellAddress = shape->GetLinkedCell()->GetRangeAddress();

// Get the shape with the name "NewShapes" from the collection
shape = prstGeomShapeCollection->Get(L"NewShapes");

// Get the cell address linked to the shape
cellAddress = shape->GetLinkedCell()->GetRangeAddress();
```

---

# Hide or Unhide Shape in Excel
## Demonstrates how to hide or unhide shapes in an Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Hide the second shape in the worksheet
sheet->GetPrstGeomShapes()->Get(1)->SetVisible(false);
```

---

# spire.xls cpp shapes
## insert various shapes into Excel sheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add a triangle shape.
intrusive_ptr<IPrstGeomShape> triangle = sheet->GetPrstGeomShapes()->AddPrstGeomShape(2, 2, 100, 100, PrstGeomShapeType::Triangle);
//Fill the triangle with solid color.
triangle->GetFill()->SetForeColor(Color::GetYellow());
triangle->GetFill()->SetFillType(ShapeFillType::SolidColor);

//Add a heart shape.
intrusive_ptr<IPrstGeomShape> heart = sheet->GetPrstGeomShapes()->AddPrstGeomShape(2, 5, 100, 100, PrstGeomShapeType::Heart);
//Fill the heart with gradient color.
heart->GetFill()->SetForeColor(Color::GetRed());
heart->GetFill()->SetFillType(ShapeFillType::Gradient);

//Add an arrow shape with default color.
intrusive_ptr<IPrstGeomShape> arrow = sheet->GetPrstGeomShapes()->AddPrstGeomShape(10, 2, 100, 100, PrstGeomShapeType::CurvedRightArrow);

//Add a cloud shape.
intrusive_ptr<IPrstGeomShape> cloud = sheet->GetPrstGeomShapes()->AddPrstGeomShape(10, 5, 100, 100, PrstGeomShapeType::Cloud);
//Fill the cloud with custom picture
cloud->GetFill()->CustomPicture(new Stream(inputImg.c_str()), L"SpireXls.png");
cloud->GetFill()->SetFillType(ShapeFillType::Picture);
```

---

# spire.xls cpp shape shadow
## modify shadow style for shape
```cpp
//Get the third shape from the worksheet.
intrusive_ptr<IPrstGeomShape> shape = sheet->GetPrstGeomShapes()->Get(2);

//Set the shadow style for the shape.
shape->GetShadow()->SetAngle(90);
shape->GetShadow()->SetTransparency(30);
shape->GetShadow()->SetDistance(10);
shape->GetShadow()->SetSize(130);
shape->GetShadow()->SetColor(Spire::Xls::Color::GetYellow());
shape->GetShadow()->SetBlur(30);
shape->GetShadow()->SetHasCustomStyle(true);
```

---

# spire.xls cpp shape shadow
## set shadow style for shape
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add an ellipse shape.
intrusive_ptr<IPrstGeomShape> ellipse = sheet->GetPrstGeomShapes()->AddPrstGeomShape(5, 5, 150, 100, PrstGeomShapeType::Ellipse);

//Set the shadow style for the ellipse.
ellipse->GetShadow()->SetAngle(90);
ellipse->GetShadow()->SetDistance(10);
ellipse->GetShadow()->SetSize(150);
ellipse->GetShadow()->SetColor(Spire::Xls::Color::GetGray());
ellipse->GetShadow()->SetBlur(30);
ellipse->GetShadow()->SetTransparency(1);
ellipse->GetShadow()->SetHasCustomStyle(true);
```

---

# spire.xls cpp shapes
## set shape order in Excel worksheets
```cpp
//Bring the picture forward one level
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetPictures()->Get(0)->ChangeLayer(ShapeLayerChangeType::BringForward);

//Bring the image in front of all other objects
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetPictures()->Get(0)->ChangeLayer(ShapeLayerChangeType::BringToFront);

//Send the shape back one level
intrusive_ptr<XlsShape> shape = dynamic_pointer_cast<XlsShape>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(2))->GetPrstGeomShapes()->Get(1));
shape->ChangeLayer(ShapeLayerChangeType::SendBackward);

//Send the shape behind all other objects
shape = dynamic_pointer_cast<XlsShape>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(3))->GetPrstGeomShapes()->Get(1));
shape->ChangeLayer(ShapeLayerChangeType::SendToBack);
```

---

# spire.xls cpp shape conversion
## convert Excel shape to image
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first shape from the first worksheet
intrusive_ptr<XlsShape> shape = dynamic_pointer_cast<XlsShape>(sheet->GetPrstGeomShapes()->Get(0));

//Save the shape to a image
intrusive_ptr<Stream> img = shape->SaveToImage();
img->Save(outputFile.c_str());
img->Close();

workbook->Dispose();
```

---

# c++ excel shape texture
## tile picture as texture in shape
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first shape
intrusive_ptr<IPrstGeomShape> shape = sheet->GetPrstGeomShapes()->Get(0);

//Fill shape with texture
shape->GetFill()->SetFillType(ShapeFillType::Texture);

//Custom texture with picture
shape->GetFill()->CustomTexture(L"Logo.png");

//Tile pciture as texture 
shape->GetFill()->SetTile(true);
```

---

# spire.xls cpp conditional formatting
## apply color scales to data range
```cpp
//Add color scales.
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
xcfs->AddRange(sheet->GetAllocatedRange());
intrusive_ptr<IConditionalFormat> format = xcfs->AddCondition();
format->SetFormatType(ConditionalFormatType::ColorScale);
```

---

# spire.xls cpp conditional formatting
## apply conditional formatting to excel cells
```cpp
//Create conditional formatting rule.
intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
xcfs1->AddRange(sheet->GetAllocatedRange());
intrusive_ptr<IConditionalFormat> format1 = xcfs1->AddCondition();
format1->SetFormatType(ConditionalFormatType::CellValue);
format1->SetFirstFormula(L"800");
format1->SetOperator(ComparisonOperatorType::Greater);
format1->SetFontColor(Spire::Xls::Color::GetRed());
format1->SetBackColor(Spire::Xls::Color::GetLightSalmon());

//Create conditional formatting rule.
intrusive_ptr<XlsConditionalFormats> xcfs2 = sheet->GetConditionalFormats()->Add();
xcfs2->AddRange(sheet->GetAllocatedRange());
intrusive_ptr<IConditionalFormat> format2 = xcfs1->AddCondition();
format2->SetFormatType(ConditionalFormatType::CellValue);
format2->SetFirstFormula(L"300");
format2->SetOperator(ComparisonOperatorType::Less);
format2->SetFontColor(Spire::Xls::Color::GetGreen());
format2->SetBackColor(Spire::Xls::Color::GetLightBlue());
```

---

# spire.xls cpp conditional formatting
## apply data bars to cell range
```cpp
//Add data bars.
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
xcfs->AddRange(sheet->GetAllocatedRange());
intrusive_ptr<IConditionalFormat> format = xcfs->AddCondition();
format->SetFormatType(ConditionalFormatType::DataBar);
format->GetDataBar()->SetBarColor(Spire::Xls::Color::GetCadetBlue());
```

---

# c++ excel gradient fill
## apply gradient fill effects to excel cells
```cpp
//Get "B5" cell
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"));
//Set row height and column width
range->SetRowHeight(50);
range->SetColumnWidth(30);
range->SetText(L"Hello");

//Set alignment style
range->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

//Set gradient filling effects
range->GetStyle()->GetInterior()->SetFillPattern(ExcelPatternType::Gradient);
range->GetStyle()->GetInterior()->GetGradient()->SetForeColor(Spire::Xls::Color::FromArgb(255, 255, 255));
range->GetStyle()->GetInterior()->GetGradient()->SetBackColor(Spire::Xls::Color::FromArgb(79, 129, 189));
range->GetStyle()->GetInterior()->GetGradient()->TwoColorGradient(GradientStyleType::Horizontal, GradientVariantsType::ShadingVariants1);
```

---

# spire.xls cpp formatting
## apply icon sets to cell range
```cpp
//Add icon sets.
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
xcfs->AddRange(sheet->GetAllocatedRange());
intrusive_ptr<IConditionalFormat> format = xcfs->AddCondition();
format->SetFormatType(ConditionalFormatType::IconSet);
format->GetIconSet()->SetIconSetType(IconSetType::ThreeTrafficLights1);
```

---

# spire.xls cpp colors and palette
## working with custom colors in excel palette and applying to cells
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Adding Orchid color to the palette at 60th index
workbook->ChangePaletteColor(Spire::Xls::Color::GetOrchid(), 60);

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"));
cell->SetText(L"Welcome to use Spire.XLS");

//Set the Orchid (custom) color to the font
cell->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetOrchid());
cell->GetStyle()->GetFont()->SetSize(20);
cell->AutoFitColumns();
cell->AutoFitRows();
```

---

# spire.xls cpp conditional formatting
## apply conditional formatting rules to Excel cells
```cpp
//Create conditional formatting rule
intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
xcfs1->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D1")));
intrusive_ptr<IConditionalFormat> cf1 = xcfs1->AddCondition();
cf1->SetFormatType(ConditionalFormatType::CellValue);
cf1->SetFirstFormula(L"150");
cf1->SetOperator(ComparisonOperatorType::Greater);
cf1->SetFontColor(Spire::Xls::Color::GetRed());
cf1->SetBackColor(Spire::Xls::Color::GetLightBlue());

intrusive_ptr<XlsConditionalFormats> xcfs2 = sheet->GetConditionalFormats()->Add();
xcfs2->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:D2")));
intrusive_ptr<IConditionalFormat> cf2 = xcfs2->AddCondition();
cf2->SetFormatType(ConditionalFormatType::CellValue);
cf2->SetFirstFormula(L"500");
cf2->SetOperator(ComparisonOperatorType::Less);
//Set border color
cf2->SetLeftBorderColor(Spire::Xls::Color::GetPink());
cf2->SetRightBorderColor(Spire::Xls::Color::GetPink());
cf2->SetTopBorderColor(Spire::Xls::Color::GetDeepSkyBlue());
cf2->SetBottomBorderColor(Spire::Xls::Color::GetDeepSkyBlue());
cf2->SetLeftBorderStyle(LineStyleType::Medium);
cf2->SetRightBorderStyle(LineStyleType::Thick);
cf2->SetTopBorderStyle(LineStyleType::Double);
cf2->SetBottomBorderStyle(LineStyleType::Double);

//Create conditional formatting rule
intrusive_ptr<XlsConditionalFormats> xcfs3 = sheet->GetConditionalFormats()->Add();
xcfs3->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3:D3")));
intrusive_ptr<IConditionalFormat> cf3 = xcfs3->AddCondition();
cf3->SetFormatType(ConditionalFormatType::CellValue);
cf3->SetFirstFormula(L"300");
cf3->SetSecondFormula(L"500");
cf3->SetOperator(ComparisonOperatorType::Between);
cf3->SetBackColor(Spire::Xls::Color::GetYellow());

//Create conditional formatting rule
intrusive_ptr<XlsConditionalFormats> xcfs4 = sheet->GetConditionalFormats()->Add();
xcfs4->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4:D4")));
intrusive_ptr<IConditionalFormat> cf4 = xcfs4->AddCondition();
cf4->SetFormatType(ConditionalFormatType::CellValue);
cf4->SetFirstFormula(L"100");
cf4->SetSecondFormula(L"200");
cf4->SetOperator(ComparisonOperatorType::NotBetween);
//Set fill pattern type
cf4->SetFillPattern(ExcelPatternType::ReverseDiagonalStripe);
//Set foreground color
cf4->SetColor(Spire::Xls::Color::FromArgb(255, 255, 0));
//Set background color
cf4->SetBackColor(Spire::Xls::Color::FromArgb(0, 255, 255));
```

---

# spire.xls cpp conditional formatting
## conditionally format dates in excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Highlight cells that contain a date occurring in the last 7 days.
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
xcfs->AddRange(sheet->GetAllocatedRange());
intrusive_ptr<IConditionalFormat> conditionalFormat = xcfs->AddTimePeriodCondition(TimePeriodType::Last7Days);
conditionalFormat->SetBackColor(Spire::Xls::Color::GetOrange());
```

---

# spire.xls cpp conditional formatting
## create formula-based conditional formatting
```cpp
//Set the conditional formatting formula and apply the rule to the chosen cell range.
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
xcfs->AddRange(range);
intrusive_ptr<IConditionalFormat> conditional = xcfs->AddCondition();
conditional->SetFormatType(ConditionalFormatType::Formula);
conditional->SetFirstFormula(L"=($A1<$B1)");
conditional->SetBackKnownColor(ExcelColors::Yellow);
```

---

# spire.xls cpp font styles
## apply various font styles to Excel cells
```cpp
//Set font style
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->GetFont()->SetFontName(L"Comic Sans MS");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D2"))->GetStyle()->GetFont()->SetFontName(L"Corbel");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:D7"))->GetStyle()->GetFont()->SetFontName(L"Aleo");

//Set font size
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->GetFont()->SetSize(45);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D3"))->GetStyle()->GetFont()->SetSize(25);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:D7"))->GetStyle()->GetFont()->SetSize(12);

//Set excel cell data to be bold
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D2"))->GetStyle()->GetFont()->SetIsBold(true);

//Set excel cell data to be underline
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:B7"))->GetStyle()->GetFont()->SetUnderline(FontUnderlineType::Single);

//set excel cell data color
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetCornflowerBlue());
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D2"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetCadetBlue());
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:D7"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetFirebrick());

//set excel cell data to be italic
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:D7"))->GetStyle()->GetFont()->SetIsItalic(true);

//Add strikethrough
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D3"))->GetStyle()->GetFont()->SetIsStrikethrough(true);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D7"))->GetStyle()->GetFont()->SetIsStrikethrough(true);
```

---

# Spire.XLS C++ Cell Formatting
## Setting foreground and background colors for Excel cells
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();
workbook->SetVersion(ExcelVersion::Version2010);

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create a new style
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle1");

//Set filling pattern type
style->GetInterior()->SetFillPattern(ExcelPatternType::Gradient);

//Set filling Background color
style->GetInterior()->GetGradient()->SetBackKnownColor(ExcelColors::Green);

//Set filling Foreground color
style->GetInterior()->GetGradient()->SetForeKnownColor(ExcelColors::Yellow);

style->GetInterior()->GetGradient()->SetGradientStyle(GradientStyleType::From_Center);

//Apply the style to "B2" cell
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetCellStyleName(style->GetName());
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetText(L"Test");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetRowHeight(30);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetColumnWidth(50);


//Create a new style
style = workbook->GetStyles()->Add(L"newStyle2");

//Set filling pattern type
style->GetInterior()->SetFillPattern(ExcelPatternType::Gradient);

//Set filling Foreground color
style->GetInterior()->GetGradient()->SetForeKnownColor(ExcelColors::Red);

//Apply the style to "B4" cell
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetCellStyleName(style->GetName());
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetRowHeight(30);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetColumnWidth(60);
```

---

# spire.xls cpp column formatting
## format a column with custom style
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create a new style
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");

//Set the vertical alignment of the text
style->SetVerticalAlignment(VerticalAlignType::Center);

//Set the horizontal alignment of the text
style->SetHorizontalAlignment(HorizontalAlignType::Center);

//Set the font color of the text
style->GetFont()->SetColor(Spire::Xls::Color::GetBlue());

//Shrink the text to fit in the cell
style->SetShrinkToFit(true);

//Set the bottom border color of the cell to OrangeRed
style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Spire::Xls::Color::GetOrangeRed());

//Set the bottom border type of the cell to Dotted
style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Dotted);

//Apply the style to the first column
sheet->GetColumns()->GetItem(0)->SetCellStyleName(style->GetName());

sheet->GetColumns()->GetItem(0)->SetText(L"Test");
```

---

# spire xls cpp row formatting
## format a row in excel with custom style
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create a new style
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");

//Set the vertical alignment of the text
style->SetVerticalAlignment(VerticalAlignType::Center);

//Set the horizontal alignment of the text
style->SetHorizontalAlignment(HorizontalAlignType::Center);

//Set the font color of the text
style->GetFont()->SetColor(Spire::Xls::Color::GetBlue());

//Shrink the text to fit in the cell
style->SetShrinkToFit(true);

//Set the bottom border color of the cell to OrangeRed
style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Spire::Xls::Color::GetOrangeRed());

//Set the bottom border type of the cell to Dotted
style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Dotted);

//Apply the style to the second row
sheet->GetRows()->GetItem(1)->SetCellStyleName(style->GetName());

sheet->GetRows()->GetItem(1)->SetText(L"Test");
```

---

# spire.xls cpp style formatting
## create and apply cell style to range
```cpp
//Create a style
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");
//Set the shading color
style->SetColor(Spire::Xls::Color::GetDarkGray());
//Set the font color
style->GetFont()->SetColor(Spire::Xls::Color::GetWhite());
//Set font name
style->GetFont()->SetFontName(L"Times New Roman");
//Set font size
style->GetFont()->SetSize(12);
//Set bold for the font
style->GetFont()->SetIsBold(true);
//Set text rotation
style->SetRotation(45);
//Set alignment
style->SetHorizontalAlignment(HorizontalAlignType::Center);
style->SetVerticalAlignment(VerticalAlignType::Center);

//Set the style for the specific range
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetRange(L"A1:J1")->SetCellStyleName(style->GetName());
```

---

# spire.xls cpp color
## get ARGB color data from Excel cells
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get font color
intrusive_ptr<Spire::Xls::Color> color1 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->GetStyle()->GetFont()->GetColor();

//Read ARGB data of Color
int a, r, g, b;
a = color1->GetA(), r = color1->GetR(), g = color1->GetG(), b = color1->GetB();
wstring argb = color1->ToString();

intrusive_ptr<Spire::Xls::Color> color2 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->GetStyle()->GetFont()->GetColor();
a = color2->GetA(), r = color2->GetR(), g = color2->GetG(), b = color2->GetB();
argb = color2->ToString();

intrusive_ptr<Spire::Xls::Color> color3 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->GetStyle()->GetFont()->GetColor();
a = color3->GetA(), r = color3->GetR(), g = color3->GetG(), b = color3->GetB();
argb = color2->ToString();
```

---

# spire.xls cpp style manipulation
## get and set cell style with font formatting
```cpp
//Get "B4" cell
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"));
//Get the style of cell
intrusive_ptr<CellStyle> style = dynamic_pointer_cast<CellStyle>(range->GetStyle());
style->GetFont()->SetFontName(L"Calibri");
style->GetFont()->SetIsBold(true);
style->GetFont()->SetSize(15);
style->GetFont()->SetColor(Spire::Xls::Color::GetCornflowerBlue());

range->SetStyle(style);
```

---

# spire.xls cpp conditional formatting
## highlight cells above and below average values
```cpp
//Add conditional format.
intrusive_ptr<XlsConditionalFormats> format1 = sheet->GetConditionalFormats()->Add();
//Set the cell range to apply the formatting.
format1->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E2:E10")));
//Add below average condition.
intrusive_ptr<IConditionalFormat> cf1 = format1->AddAverageCondition(AverageType::Below);
//Highlight cells below average values.
cf1->SetBackColor(Spire::Xls::Color::GetSkyBlue());

//Add conditional format.
intrusive_ptr<XlsConditionalFormats> format2 = sheet->GetConditionalFormats()->Add();
//Set the cell range to apply the formatting.
format2->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E2:E10")));
//Add above average condition.
intrusive_ptr<IConditionalFormat> cf2 = format1->AddAverageCondition(AverageType::Above);
//Highlight cells above average values.
cf2->SetBackColor(Spire::Xls::Color::GetOrange());
```

---

# spire.xls cpp conditional formatting
## highlight duplicate and unique values in Excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Use conditional formatting to highlight duplicate values in range "C2:C10" with IndianRed color.
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
xcfs->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:C10")));
intrusive_ptr<IConditionalFormat> format1 = xcfs->AddCondition();
format1->SetFormatType(ConditionalFormatType::DuplicateValues);
format1->SetBackColor(Spire::Xls::Color::GetIndianRed());

//Use conditional formatting to highlight unique values in range "C2:C10" with Yellow color.
intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
xcfs1->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:C10")));
intrusive_ptr<IConditionalFormat> format2 = xcfs->AddCondition();
format2->SetFormatType(ConditionalFormatType::UniqueValues);
format2->SetBackColor(Spire::Xls::Color::GetYellow());
```

---

# spire.xls cpp conditional formatting
## highlight top and bottom ranked values in excel
```cpp
//Apply conditional formatting to range D2:D10 to highlight the top 2 values.
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
xcfs->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2:D10")));
intrusive_ptr<IConditionalFormat> format1 = xcfs->AddTopBottomCondition(TopBottomType::Top, 2);
format1->SetFormatType(ConditionalFormatType::TopBottom);
format1->SetBackColor(Spire::Xls::Color::GetRed());

//Apply conditional formatting to range E2:E10 to highlight the bottom 2 values.
intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
xcfs1->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E2:E10")));
intrusive_ptr<IConditionalFormat> format2 = xcfs1->AddTopBottomCondition(TopBottomType::Bottom, 2);
format2->SetFormatType(ConditionalFormatType::TopBottom);
format2->SetBackColor(Spire::Xls::Color::GetForestGreen());
```

---

# spire.xls cpp formatting
## set cell text indentation level
```cpp
//Access the "B5" cell from the worksheet
intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"));

//Add some value to the "B5" cell
cell->SetText(L"Hello Spire!");

//Set the indentation level of the text (inside the cell) to 2
cell->GetStyle()->SetIndentLevel(2);
```

---

# spire.xls cpp cell activation
## Make cell active in Excel worksheet
```cpp
//Get the 2nd sheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

//Set the 2nd sheet as an active sheet.
sheet->Activate();

//Set B2 cell as an active cell in the worksheet.
sheet->SetActiveCell(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2")));

//Set the B column as the first visible column in the worksheet.
sheet->SetFirstVisibleColumn(1);

//Set the 2nd row as the first visible row in the worksheet.
sheet->SetFirstVisibleRow(1);
```

---

# spire.xls cpp number formatting
## apply various number formats to Excel cells
```cpp
//Set title for number formatting section
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10"))->SetText(L"NUMBER FORMATTING");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10"))->GetStyle()->GetFont()->SetIsBold(true);

//Apply different number formats to cells
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B13"))->SetText(L"0");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C13"))->SetNumberValue(1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C13"))->SetNumberFormat(L"0");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B14"))->SetText(L"0.00");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C14"))->SetNumberValue(1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C14"))->SetNumberFormat(L"0.00");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B15"))->SetText(L"#,##0.00");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C15"))->SetNumberValue(1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C15"))->SetNumberFormat(L"#,##0.00");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B16"))->SetText(L"$#,##0.00");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C16"))->SetNumberValue(1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C16"))->SetNumberFormat(L"$#,##0.00");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B17"))->SetText(L"0;[Red]-0");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C17"))->SetNumberValue(-1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C17"))->SetNumberFormat(L"0;[Red]-0");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B18"))->SetText(L"0.00;[Red]-0.00");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C18"))->SetNumberValue(-1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C18"))->SetNumberFormat(L"0.00;[Red]-0.00");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B19"))->SetText(L"#,##0;[Red]-#,##0");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C19"))->SetNumberValue(-1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C19"))->SetNumberFormat(L"#,##0;[Red]-#,##0");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B20"))->SetText(L"#,##0.00;[Red]-#,##0.000");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C20"))->SetNumberValue(-1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C20"))->SetNumberFormat(L"#,##0.00;[Red]-#,##0.00");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B21"))->SetText(L"0.00E+00");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C21"))->SetNumberValue(1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C21"))->SetNumberFormat(L"0.00E+00");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B22"))->SetText(L"0.00%");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C22"))->SetNumberValue(1234.5678);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C22"))->SetNumberFormat(L"0.00%");

//Apply background color to format description cells
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B13:B22"))->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

//AutoFit columns for better visibility
sheet->AutoFitColumn(2);
sheet->AutoFitColumn(3);
```

---

# spire.xls cpp border
## set border style for excel cells
```cpp
//Get the cell range where you want to apply border style
intrusive_ptr<CellRange> cr = dynamic_pointer_cast<CellRange>(sheet->GetRange(sheet->GetFirstRow(), sheet->GetFirstColumn(), sheet->GetLastRow(), sheet->GetLastColumn()));

//Apply border style 
cr->GetBorders()->SetLineStyle(LineStyleType::Double);
intrusive_ptr<IBorder> border = cr->GetBorders()->Get(BordersLineType::DiagonalDown);
border->SetLineStyle(LineStyleType::None);
cr->GetBorders()->Get(BordersLineType::DiagonalUp)->SetLineStyle(LineStyleType::None);
cr->GetBorders()->SetColor(Spire::Xls::Color::GetCadetBlue());
```

---

# spire.xls cpp databar
## set border to data bar in conditional formatting
```cpp
// Get existing conditional format and set data bar border
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Get(0); 
intrusive_ptr<IConditionalFormat>cf = xcfs->Get(0);
intrusive_ptr<Spire::Xls::DataBar> dataBar1 = cf->GetDataBar();
dataBar1->GetBarBorder()->SetType(DataBarBorderType::DataBarBorderSolid);
dataBar1->GetBarBorder()->SetColor(Color::GetRed());

// Create new conditional format with data bar border
intrusive_ptr<XlsConditionalFormats>xcfs2 = sheet->GetConditionalFormats()->Add(); 
intrusive_ptr<IConditionalFormat>cf2 = xcfs2->AddCondition(); 
cf2->SetFormatType(ConditionalFormatType::DataBar); 
cf2->GetDataBar()->GetBarBorder()->SetType(DataBarBorderType::DataBarBorderSolid); 
cf2->GetDataBar()->GetBarBorder()->SetColor(Color::GetRed()); 
cf2->GetDataBar()->SetBarColor(Color::GetGreenYellow());
```

---

# spire.xls cpp conditional formatting
## set conditional format formula for cells
```cpp
//Add ConditionalFormat
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();

//Define the range
xcfs->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5")));

//Add condition
intrusive_ptr<IConditionalFormat> format = xcfs->AddCondition();
format->SetFormatType(ConditionalFormatType::CellValue);

//If greater than 1000
format->SetFirstFormula(L"1000");
format->SetOperator(ComparisonOperatorType::Greater);
format->SetBackColor(Spire::Xls::Color::GetOrange());
```

---

# c++ conditional formatting
## set row color by conditional format
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Select the range that you want to format.
intrusive_ptr<CellRange> dataRange = dynamic_pointer_cast<CellRange>(sheet->GetAllocatedRange());

//Set conditional formatting.
intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
xcfs->AddRange(dataRange);
intrusive_ptr<IConditionalFormat> format1 = xcfs->AddCondition();
//Determines the cells to format.
format1->SetFirstFormula(L"=MOD(ROW(),2)=0");
//Set conditional formatting type
format1->SetFormatType(ConditionalFormatType::Formula);
//Set the color.
format1->SetBackColor(Spire::Xls::Color::GetLightSeaGreen());

//Set the backcolor of the odd rows as Yellow.
intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
xcfs1->AddRange(dataRange);
intrusive_ptr<IConditionalFormat> format2 = xcfs->AddCondition();
format2->SetFirstFormula(L"=MOD(ROW(),2)=1");
format2->SetFormatType(ConditionalFormatType::Formula);
format2->SetBackColor(Spire::Xls::Color::GetYellow());
```

---

# spire.xls cpp formatting
## set traffic lights icons in Excel cells
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the height of row and width of column for Excel cell range.
sheet->GetAllocatedRange()->SetRowHeight(20);
sheet->GetAllocatedRange()->SetColumnWidth(25);

//Add a conditional formatting.
intrusive_ptr<XlsConditionalFormats> conditional = sheet->GetConditionalFormats()->Add();
conditional->AddRange(sheet->GetAllocatedRange());
intrusive_ptr<IConditionalFormat> format1 = conditional->AddCondition();

//Add a conditional formatting of cell range and set its type to CellValue.
format1->SetFormatType(ConditionalFormatType::CellValue);
format1->SetFirstFormula(L"300");
format1->SetOperator(ComparisonOperatorType::Less);
format1->SetFontColor(Spire::Xls::Color::GetBlack());
format1->SetBackColor(Spire::Xls::Color::GetLightSkyBlue());

//Add a conditional formatting of cell range and set its type to IconSet.
conditional->AddRange(sheet->GetAllocatedRange());
intrusive_ptr<IConditionalFormat> format = conditional->AddCondition();
format->SetFormatType(ConditionalFormatType::IconSet);
format->GetIconSet()->SetIconSetType(IconSetType::ThreeTrafficLights1);
```

---

# spire.xls cpp conditional formatting
## apply simple conditional formatting to Excel cells
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

sheet->GetAllocatedRange()->SetRowHeight(15);
sheet->GetAllocatedRange()->SetColumnWidth(16);

//Create conditional formatting rule
intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
xcfs1->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D1")));
intrusive_ptr<IConditionalFormat> cf1 = xcfs1->AddCondition();
cf1->SetFormatType(ConditionalFormatType::CellValue);
cf1->SetFirstFormula(L"150");
cf1->SetOperator(ComparisonOperatorType::Greater);
cf1->SetFontColor(Spire::Xls::Color::GetRed());
cf1->SetBackColor(Spire::Xls::Color::GetLightBlue());

intrusive_ptr<XlsConditionalFormats> xcfs2 = sheet->GetConditionalFormats()->Add();
xcfs2->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:D2")));
intrusive_ptr<IConditionalFormat> cf2 = xcfs2->AddCondition();
cf2->SetFormatType(ConditionalFormatType::CellValue);
cf2->SetFirstFormula(L"300");
cf2->SetOperator(ComparisonOperatorType::Less);
//Set border color
cf2->SetLeftBorderColor(Spire::Xls::Color::GetPink());
cf2->SetRightBorderColor(Spire::Xls::Color::GetPink());
cf2->SetTopBorderColor(Spire::Xls::Color::GetDeepSkyBlue());
cf2->SetBottomBorderColor(Spire::Xls::Color::GetDeepSkyBlue());
cf2->SetLeftBorderStyle(LineStyleType::Medium);
cf2->SetRightBorderStyle(LineStyleType::Thick);
cf2->SetTopBorderStyle(LineStyleType::Double);
cf2->SetBottomBorderStyle(LineStyleType::Double);

//Add data bars
intrusive_ptr<XlsConditionalFormats> xcfs3 = sheet->GetConditionalFormats()->Add();
xcfs3->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3:D3")));
intrusive_ptr<IConditionalFormat> cf3 = xcfs3->AddCondition();
cf3->SetFormatType(ConditionalFormatType::DataBar);
cf3->GetDataBar()->SetBarColor(Spire::Xls::Color::GetCadetBlue());

//Add icon sets
intrusive_ptr<XlsConditionalFormats> xcfs4 = sheet->GetConditionalFormats()->Add();
xcfs4->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4:D4")));
intrusive_ptr<IConditionalFormat> cf4 = xcfs4->AddCondition();
cf4->SetFormatType(ConditionalFormatType::IconSet);
cf4->GetIconSet()->SetIconSetType(IconSetType::ThreeTrafficLights1);

//Add color scales
intrusive_ptr<XlsConditionalFormats> xcfs5 = sheet->GetConditionalFormats()->Add();
xcfs5->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:D5")));
intrusive_ptr<IConditionalFormat> cf5 = xcfs5->AddCondition();
cf5->SetFormatType(ConditionalFormatType::ColorScale);

//Highlight duplicate values in range "A6:D6" with BurlyWood color
intrusive_ptr<XlsConditionalFormats> xcfs6 = sheet->GetConditionalFormats()->Add();
xcfs6->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6:D6")));
intrusive_ptr<IConditionalFormat> cf6 = xcfs6->AddCondition();
cf6->SetFormatType(ConditionalFormatType::DuplicateValues);
cf6->SetBackColor(Spire::Xls::Color::GetBurlyWood());
```

---

# spire.xls cpp text alignment
## set cell text alignment and rotation in Excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the vertical alignment to Top
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1:C1"))->GetStyle()->SetVerticalAlignment(VerticalAlignType::Top);

//Set the vertical alignment to Center
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:C2"))->GetStyle()->SetVerticalAlignment(VerticalAlignType::Center);

//Set the vertical alignment of to Bottom
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:C3"))->GetStyle()->SetVerticalAlignment(VerticalAlignType::Bottom);

//Set the horizontal alignment to General
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4:C4"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::General);

//Set the horizontal alignment of to Left
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5:C5"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Left);

//Set the horizontal alignment of to Center
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B6:C6"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

//Set the horizontal alignment of to Right
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B7:C7"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Right);

//Set the rotation degree
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8:C8"))->GetStyle()->SetRotation(45);

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B9:C9"))->GetStyle()->SetRotation(90);

//Set the row height of cell
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8:C9"))->SetRowHeight(60);
```

---

# Excel Cell Text Direction
## Set text direction in Excel cells
```cpp
// Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

// Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Access the "B5" cell from the worksheet
intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"));

// Add some value to the "B5" cell
cell->SetText(L"Hello Spire!");

// Set the reading order from right to left of the text in the "B5" cell
cell->GetStyle()->SetReadingOrder(ReadingOrderType::RightToLeft);
```

---

# spire xls cpp predefined styles
## create and apply predefined styles to cells
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Create a new style
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");
style->GetFont()->SetFontName(L"Calibri");
style->GetFont()->SetIsBold(true);
style->GetFont()->SetSize(15);
style->GetFont()->SetColor(Spire::Xls::Color::GetCornflowerBlue());

//Get "B5" cell
intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"));
range->SetText(L"Welcome to use Spire.XLS");
range->SetCellStyleName(style->GetName());
range->AutoFitColumns();
```

---

# spire.xls cpp style
## create and apply style to excel cells
```cpp
//Create a new style
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");

//Set the vertical alignment of the text in the cell
style->SetVerticalAlignment(VerticalAlignType::Center);

//Set the horizontal alignment of the text in the cell
style->SetHorizontalAlignment(HorizontalAlignType::Center);

//Set the font color of the text in the cell
style->GetFont()->SetColor(Spire::Xls::Color::GetBlue());

//Shrink the text to fit in the cell
style->SetShrinkToFit(true);

//Set the bottom border color of the cell to GreenYellow
style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Spire::Xls::Color::GetGreenYellow());

//Set the bottom border type of the cell to Medium
style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Medium);

//Assign the Style object to the cell
cell->SetStyle(style);

//Apply the same style to some other cells
sheet->GetRange(L"B4")->SetStyle(style);
sheet->GetRange(L"C3")->SetCellStyleName(style->GetName());
sheet->GetRange(L"D4")->SetStyle(style);
```

---

# spire.xls cpp named range formula
## insert formula with named range in Excel
```cpp
//Create a named range
intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Add(L"NewNamedRange");

NamedRange->SetNameLocal(L"=SUM(A1+A2)");

//Set the formula
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetFormula(L"NewNamedRange");
```

---

# spire.xls cpp read formulas
## read formula and its calculated value from an Excel cell
```cpp
// Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

// Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

// Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

wstring formula = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C14"))->GetFormula();
wstring value = to_wstring(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C14"))->GetFormulaNumberValue());
```

---

# spire xls cpp addin functions
## register and use addin functions in excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Register AddIn function
workbook->GetAddInFunctions()->Add(inputFile.c_str(), L"TEST_UDF");
workbook->GetAddInFunctions()->Add(inputFile.c_str(), L"TEST_UDF1");
//Get the first sheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Call AddIn function
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetFormula(L"=TEST_UDF()");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetFormula(L"=TEST_UDF1()");
```

---

# spire.xls cpp subtotal formula
## implement SUBTOTAL formulas in Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add SUBTOTAL formulas
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetFormula(L"=SUBTOTAL(1,A1:C3)");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetFormula(L"=SUBTOTAL(2,A1:C3)");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetFormula(L"=SUBTOTAL(5,A1:C3)");

//Calculate Formulas
workbook->CalculateAllValue();
```

---

# spire.xls cpp array formulas
## using array formulas in Excel
```cpp
//Write array formula
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:C6"))->SetFormulaArray(L"=LINEST(A1:A3,B1:C3,TRUE,TRUE)");

//Calculate Formulas
workbook->CalculateAllValue();
```

---

# spire.xls cpp r1c1 formula
## use array R1C1 formula in Excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set some sample data
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetNumberValue(1);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetNumberValue(2);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetNumberValue(3);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetNumberValue(4);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(5);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(6);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetNumberValue(7);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetNumberValue(8);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetNumberValue(9);

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetText(L"Sum:");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Right);

//Write array R1C1 formula
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetFormulaArrayR1C1(L"=SUM(R[-3]C[-2]:R[-1]C)");

//Calculate Formulas
workbook->CalculateAllValue();
```

---

# spire.xls cpp r1c1 formula
## demonstrate how to use R1C1-style formulas in Excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetText(L"Sum:");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Right);
//Write array R1C1 formula
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetFormulaR1C1(L"=SUM(R[-3]C[-2]:R[-1]C)");

//Calculate Formulas
workbook->CalculateAllValue();
```

---

# spire.xls cpp formulas
## write various Excel formulas in cells
```cpp
// String formula
currentFormula = (L"=\"hello\"");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(L"=\"hello\"");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());
wstring s(L"test");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 3))->SetFormula((L"=\"" + s + L"\"").c_str());

// Integer formula
currentFormula = (L"=300");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

// Float formula
currentFormula = (L"=3389.639421");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

// Boolean formula
currentFormula = (L"=false");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

// Arithmetic operations
currentFormula = (L"=1+2+3+4+5-6-7+8-9");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=33*3/4-2+10");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

// Sheet reference
currentFormula = (L"=Sheet1!$B$3");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

// Sheet area reference
currentFormula = (L"=AVERAGE(Sheet1!$D$3:G$3)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

// Functions
currentFormula = (L"=Count(3,5,8,10,2,34)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=NOW()");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->GetStyle()->SetNumberFormat(L"yyyy-MM-DD");

currentFormula = (L"=SECOND(11)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=MINUTE(12)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=MONTH(9)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=DAY(10)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=TIME(4,5,7)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=DATE(6,4,2)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=RAND()");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=HOUR(12)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=MOD(5,3)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=WEEKDAY(3)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=YEAR(23)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=NOT(true)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=OR(true)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=AND(TRUE)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=VALUE(30)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=LEN(\"world\")");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=MID(\"world\",4,2)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=ROUND(7,3)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=SIGN(4)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=INT(200)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=ABS(-1.21)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=LN(15)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=EXP(20)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=SQRT(40)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=PI()");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=COS(9)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=SIN(45)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=MAX(10,30)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=MIN(5,7)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=AVERAGE(12,45)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=SUM(18,29)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=IF(4,2,2)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

currentFormula = (L"=SUBTOTAL(3,Sheet1!B2:E3)");
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());
```

---

# spire.xls cpp header footer
## change font and size for header and footer
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the new font and size for the header and footer
wstring text = sheet->GetPageSetup()->GetLeftHeader();
//"Arial Unicode MS" is font name, L"18" is font size
text = L"&\"Arial Unicode MS\"&18 Header Footer Sample by Spire.XLS ";
sheet->GetPageSetup()->SetLeftHeader(text.c_str());
sheet->GetPageSetup()->SetRightFooter(text.c_str());
```

---

# spire.xls cpp header footer
## add image to header and footer in excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile("input_file_path");

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Load an image from disk
intrusive_ptr<Stream> image = new Stream(DATAPATH L"/Demo/Logo.png");

//Set the image header
sheet->GetPageSetup()->SetLeftHeaderImage(image);
sheet->GetPageSetup()->SetLeftHeader(L"&G");

//Set the image footer
sheet->GetPageSetup()->SetCenterFooterImage(image);
sheet->GetPageSetup()->SetCenterFooter(L"&G");

//Set the view mode of the sheet
sheet->SetViewMode(ViewMode::Layout);
```

---

# spire.xls cpp header footer
## set header and footer in Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set left header,"Arial Unicode MS" is font name, L"18" is font size.
sheet->GetPageSetup()->SetLeftHeader(L"&\"Arial Unicode MS\"&14 Spire.XLS for C++ ");

//Set center footer 
sheet->GetPageSetup()->SetCenterFooter(L"Footer Text");

sheet->SetViewMode(ViewMode::Layout);
```

---

# spire.xls cpp hyperlink
## add hyperlink to text in excel cells
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add url link
intrusive_ptr<HyperLink> UrlLink = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D10"))));
UrlLink->SetTextToDisplay(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D10"))->GetText());
UrlLink->SetType(HyperLinkType::Url);
UrlLink->SetAddress(L"http://en.wikipedia.org/wiki/Chicago");

//Add email link
intrusive_ptr<XlsHyperLink> MailLink = sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E10")));
MailLink->SetTextToDisplay(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E10"))->GetText());
MailLink->SetType(HyperLinkType::Url);
MailLink->SetAddress(L"mailto:Amor.Aqua@gmail.com");
```

---

# spire.xls cpp image hyperlink
## add hyperlink to an image in Excel worksheet
```cpp
//Insert an image to a specific cell
intrusive_ptr<ExcelPicture> picture = ExcelPicture::Dynamic_cast<ExcelPicture>(sheet->GetPictures()->Add(2, 1, "path_to_image.png"));
//Add a hyperlink to the image
picture->SetHyperLink(L"https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html", true);
```

---

# spire.xls cpp hyperlink
## create hyperlink to external file in Excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(1, 1));

//Add hyperlink in the range
intrusive_ptr<HyperLink> hyperlink = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(range));

//Set the link type
hyperlink->SetType(HyperLinkType::File);

//Set the display text
hyperlink->SetTextToDisplay(L"Link To External File");

//Set file address
hyperlink->SetAddress(L"SampeB_4.xlsx");
```

---

# c++ excel hyperlink
## create hyperlink to another sheet cell
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"));

//Add hyperlink in the range
intrusive_ptr<HyperLink> hyperlink = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(range));

//Set the link type
hyperlink->SetType(HyperLinkType::Workbook);

//Set the display text
hyperlink->SetTextToDisplay(L"Link to Sheet2 cell C5");

//Set the address
hyperlink->SetAddress(L"Sheet2!C5");
```

---

# spire.xls cpp hyperlink
## modify hyperlink in excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Change the values of TextToDisplay and Address property 
intrusive_ptr<IHyperLinks> links = sheet->GetHyperLinks();
links->Get(0)->SetTextToDisplay(L"Product livedemo");
links->Get(0)->SetAddress(L"https://www.e-iceblue.com/LiveDemo.html");
```

---

# spire.xls cpp hyperlinks
## read hyperlinks from excel worksheet
```cpp
// Assuming workbook is already created and loaded with Excel file
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

wstring address1 = sheet->GetHyperLinks()->Get(0)->GetAddress();
wstring address2 = sheet->GetHyperLinks()->Get(1)->GetAddress();
```

---

# spire.xls cpp hyperlinks
## remove hyperlinks from excel worksheet
```cpp
//Remove all link content
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->ClearAll();
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->ClearAll();
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->ClearAll();

//Remove hyperlink and keep link text
sheet->GetHyperLinks()->RemoveAt(0);
sheet->GetHyperLinks()->RemoveAt(0);
sheet->GetHyperLinks()->RemoveAt(0);
```

---

# spire.xls cpp hyperlinks
## retrieve external file hyperlinks from excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Retrieve external file hyperlinks.
for (int i = 0; i < sheet->GetHyperLinks()->GetCount(); i++)
{
    intrusive_ptr<HyperLink> hyperlink = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Get(i));
    wstring address = hyperlink->GetAddress();
    wstring sheetName = dynamic_pointer_cast<XlsRange>(hyperlink->GetRange())->GetWorksheetName();
    intrusive_ptr<IXLSRange> range = hyperlink->GetRange();
    int row = range->GetRow();
    int column = range->GetColumn();
}
```

---

# spire.xls cpp hyperlinks
## add hyperlinks to excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B9"))->SetText(L"Home page");
intrusive_ptr<HyperLink> hylink1 = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10"))));
hylink1->SetType(HyperLinkType::Url);
hylink1->SetAddress(L"(http://www.e-iceblue.com)");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B11"))->SetText(L"Support");
intrusive_ptr<HyperLink> hylink2 = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B12"))));
hylink2->SetType(HyperLinkType::Url);
hylink2->SetAddress(L"mailto:support@e-iceblue.com");

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B13"))->SetText(L"Forum");
intrusive_ptr<HyperLink> hylink3 = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B14"))));
hylink3->SetType(HyperLinkType::Url);
hylink3->SetAddress(L"https://www.e-iceblue.com/forum/");
```

---

# spire.xls cpp named ranges
## format cells in named ranges
```cpp
//Get specific named range by index
intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Get(0);

//Get the cell range of the named range
intrusive_ptr<IXLSRange> range = NamedRange->GetRefersToRange();

//Set color for the range
range->GetStyle()->SetColor(Spire::Xls::Color::GetYellow());

//Set the font as bold
range->GetStyle()->GetFont()->SetIsBold(true);
```

---

# spire.xls cpp named ranges
## get all named ranges from workbook
```cpp
using namespace Spire::Xls;

//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get all named range
intrusive_ptr<INameRanges> ranges = workbook->GetNameRanges();
for (int i = 0; i < ranges->GetCount(); i++)
{
    intrusive_ptr<INamedRange> nameRange = ranges->Get(i);
}
```

---

# spire.xls cpp named range
## get address of named range from Excel workbook
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get specific named range by index
intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Get(0);

//Get the address of the named range
wstring address = NamedRange->GetRefersToRange()->GetRangeAddress();
```

---

# spire.xls cpp named ranges
## get specific named range by index and name
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get specific named range by index
wstring name1 = workbook->GetNameRanges()->Get(1)->GetName();

//Get specific named range by name
wstring name2 = workbook->GetNameRanges()->Get(L"NameRange3")->GetName();
```

---

# spire.xls cpp named range
## merge cells in a named range
```cpp
//Get specific named range by index
intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Get(0);

//Get the range of the named range
intrusive_ptr<IXLSRange> range = NamedRange->GetRefersToRange();

//Merge cells
range->Merge();
```

---

# spire.xls cpp named ranges
## create and configure named ranges in Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Creating a named range
intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Add(L"NewNamedRange");
//Setting the range of the named range
NamedRange->SetRefersToRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:E12")));

workbook->Dispose();
```

---

# spire.xls remove named range
## demonstrates how to remove named ranges from an Excel workbook
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Remove the named range by index
workbook->GetNameRanges()->RemoveAt(0);

//Remove the named range by name
workbook->GetNameRanges()->Remove(L"NameRange2");
```

---

# spire.xls cpp named range
## rename named range in excel workbook
```cpp
//Rename the named range
workbook->GetNameRanges()->Get(0)->SetName(L"RenameRange");
```

---

# spire.xls cpp named range
## create a scoped named range in Excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add range name
intrusive_ptr<INamedRange> namedRange = sheet->GetNames()->Add(L"Range1");

//Define the range
namedRange->SetRefersToRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D19")));
```

---

# spire.xls cpp named range
## set formula with named range
```cpp
//Create a named range
intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Add(L"MyNamedRange");
//Refers to range
NamedRange->SetRefersToRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10:B12")));

//Set the formula of range to named range
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B13"))->SetFormula(L"=SUM(MyNamedRange)");
```

---

# spire.xls cpp ole objects
## extract and save OLE objects from Excel worksheet
```cpp
// Convert a wide string to a standard string
std::string wstring2string(const std::wstring& wstr)
{
    std::string result;
    result.reserve(wstr.size());
    for (size_t i = 0; i < wstr.size(); ++i)
    {
        result += static_cast<char>(wstr[i] & 0xFF);
    }
    return result;
}

// Write all bytes to a file
void WriteAllBytes(std::wstring filePath, std::vector<byte> data)
{
    std::ofstream outFile(wstring2string(filePath), std::ios::out | std::ofstream::binary);
    outFile.write((char*)(&data[0]), data.size() * sizeof(byte));
    outFile.close();
}

// Get the first worksheet from the workbook
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Check if the worksheet has OleObjects (embedded objects)
if (sheet->GetHasOleObjects())
{
    // Get the count of OleObjects in the worksheet
    int count = sheet->GetOleObjects()->GetCount();
    for (int i = 0; i < count; i++)
    {
        // Get the OleObject at the specified index
        auto Object = sheet->GetOleObjects()->GetItem(i);
        OleObjectType type = Object->GetObjectType();
        
        // Extract and save OLE object based on its type
        switch (type)
        {
            // Handle Word document objects
            case OleObjectType::WordDocument:
                WriteAllBytes(L"output.docx", Object->GetOleData());
                break;

            // Handle Adobe Acrobat document objects
            case OleObjectType::AdobeAcrobatDocument:
                WriteAllBytes(L"output.pdf", Object->GetOleData());
                break;

            // Handle PowerPoint slide objects
            case OleObjectType::PowerPointSlide:
                WriteAllBytes(L"output.pptx", Object->GetOleData());
                break;

            default:
                break;
        }
    }
}
```

---

# spire.xls cpp ole objects
## insert OLE objects into Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

sheet->GetRange(L"A1")->SetText(L"Here is an OLE Object.");
//insert OLE object
intrusive_ptr<Workbook> book = new Workbook();
book->LoadFromFile(inputFile.c_str());
book->GetWorksheets()->Get(0)->GetPageSetup()->SetLeftMargin(0);
book->GetWorksheets()->Get(0)->GetPageSetup()->SetRightMargin(0);
book->GetWorksheets()->Get(0)->GetPageSetup()->SetTopMargin(0);
book->GetWorksheets()->Get(0)->GetPageSetup()->SetBottomMargin(0);
intrusive_ptr<Stream> image = book->GetWorksheets()->Get(0)->ToImage(1, 1, 19, 5);
intrusive_ptr<Spire::Xls::IOleObject> oleObject = sheet->GetOleObjects()->Add(inputFile.c_str(), image, OleLinkType::Embed);

oleObject->SetLocation(sheet->GetRange(L"B4"));
oleObject->SetObjectType(OleObjectType::ExcelWorksheet);
```

---

# spire.xls cpp page setup
## get Excel paper dimensions
```cpp
using namespace Spire::Xls;

int main() {
	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

	//Get dimensions of different paper sizes
	pageSetup->SetPaperSize(PaperSizeType::A2Paper);
	double a2Width = pageSetup->GetPageWidth();
	double a2Height = pageSetup->GetPageHeight();

	pageSetup->SetPaperSize(PaperSizeType::PaperA3);
	double a3Width = pageSetup->GetPageWidth();
	double a3Height = pageSetup->GetPageHeight();

	pageSetup->SetPaperSize(PaperSizeType::PaperA4);
	double a4Width = pageSetup->GetPageWidth();
	double a4Height = pageSetup->GetPageHeight();

	pageSetup->SetPaperSize(PaperSizeType::PaperLetter);
	double letterWidth = pageSetup->GetPageWidth();
	double letterHeight = pageSetup->GetPageHeight();
}
```

---

# spire.xls cpp page setup
## set Excel page order type
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the reference of the PageSetup of the worksheet.
intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

//Set the order type of the pages to over then down.
pageSetup->SetOrder(OrderType::OverThenDown);
```

---

# spire.xls cpp page setup
## set Excel paper size to A4
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the paper size of the worksheet as A4 paper.
sheet->GetPageSetup()->SetPaperSize(PaperSizeType::PaperA4);
```

---

# spire.xls cpp page setup
## set first page number for worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the first page number of the worksheet pages.
sheet->GetPageSetup()->SetFirstPageNumber(2);
```

---

# spire.xls cpp pagesetup
## set header and footer margins
```cpp
//Get the PageSetup object of the first worksheet.
intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

//Set the margins of header and footer.
pageSetup->SetHeaderMarginInch(2);
pageSetup->SetFooterMarginInch(2);
```

---

# spire.xls cpp page setup
## set margins of excel sheet
```cpp
//Get the PageSetup object of the first worksheet.
intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

//Set bottom,left,right and top page margins.
pageSetup->SetBottomMargin(2);
pageSetup->SetLeftMargin(1);
pageSetup->SetRightMargin(1);
pageSetup->SetTopMargin(3);
```

---

# spire.xls cpp page setup
## set various printing options for Excel worksheet
```cpp
//Get the reference of the PageSetup of the worksheet.
intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

//Allow to print gridlines.
pageSetup->SetIsPrintGridlines(true);

//Allow to print row/column headings.
pageSetup->SetIsPrintHeadings(true);

//Allow to print worksheet in black & white mode.
pageSetup->SetBlackAndWhite(true);

//Allow to print comments as displayed on worksheet.
pageSetup->SetPrintComments(PrintCommentType::InPlace);

//Allow to print worksheet with draft quality.
pageSetup->SetDraft(true);

//Allow to print cell errors as N/A.
pageSetup->SetPrintErrors(PrintErrorsType::NA);
```

---

# spire.xls cpp page setup
## set page orientation to landscape
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the page orientation to Landscape. 
sheet->GetPageSetup()->SetOrientation(PageOrientationType::Landscape);
```

---

# Spire.XLS C++ Page Setup
## Set print area of Excel file
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the reference of the PageSetup of the worksheet.
intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

//Specify the cells range of the print area.
pageSetup->SetPrintArea(L"A1:E5");
```

---

# spire.xls cpp page setup
## Set print quality of Excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the print quality of the worksheet to 180 dpi.
sheet->GetPageSetup()->SetPrintQuality(180);
```

---

# spire.xls cpp page setup
## set print title rows and columns for Excel file
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

//Define column numbers A & B as title columns.
pageSetup->SetPrintTitleColumns(L"$A:$B");

//Defining row numbers 1 & 2 as title rows.
pageSetup->SetPrintTitleRows(L"$1:$2");
```

---

# spire.xls cpp page setup
## set sheet fit to page property
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

//Set the FitToPagesTall property.
sheet->GetPageSetup()->SetFitToPagesTall(1);

//Set the FitToPagesWide property.
sheet->GetPageSetup()->SetFitToPagesWide(1);
```

---

# spire.xls cpp page setup
## center worksheet on page when printing
```cpp
//Get the PageSetup object of the first page.
intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

//Set the worksheet center on page.
pageSetup->SetCenterHorizontally(true);
pageSetup->SetCenterVertically(true);
```

---

# spire.xls cpp pivot table
## change pivot table data source
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

intrusive_ptr<CellRange> Range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C15"));
intrusive_ptr<XlsPivotTable> table = dynamic_pointer_cast<XlsPivotTable>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetPivotTables()->Get(0));

//Change data source
table->ChangeDataSource(Range);
table->GetCache()->SetIsRefreshOnLoad(false);
```

---

# Clear Pivot Table Fields
## This code demonstrates how to clear all data fields in a pivot table
```cpp
//Get the worksheet containing pivot table
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"PivotTable"));

//Get pivot table from worksheet
intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));

//Clear all the data fields
pt->GetDataFields()->Clear();

pt->CalculateData();
```

---

# spire.xls cpp pivot table consolidation functions
## apply consolidation functions to pivot table data fields
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"PivotTable"));

intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));
//Apply Average consolidation function to first data field
pt->GetDataFields()->Get(0)->SetSubtotal(SubtotalTypes::Average);
//Apply Max consolidation function to second data field
pt->GetDataFields()->Get(1)->SetSubtotal(SubtotalTypes::Max);
pt->CalculateData();
```

---

# Disable Pivot Table Ribbon in Excel
## This code demonstrates how to disable the ribbon/wizard for a pivot table in an Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"PivotTable"));

intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));
//Disable ribbon for this pivot table
pt->SetEnableWizard(false);
```

---

# c++ pivot table expand collapse rows
## expand or collapse rows in excel pivot table
```cpp
//Get the data in Pivot Table.
intrusive_ptr<XlsPivotTable> pivotTable = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));

//Calculate Data.
pivotTable->CalculateData();

//Collapse the rows.
(dynamic_pointer_cast<XlsPivotField>(pivotTable->GetPivotFields()->Get(L"Vendor No")))->HideItemDetail(L"3501", true);

//Expand the rows.
(dynamic_pointer_cast<XlsPivotField>(pivotTable->GetPivotFields()->Get(L"Vendor No")))->HideItemDetail(L"3502", false);
```

---

# C++ PivotTable Data Field Formatting
## Format PivotTable data field to show as percentage of column
```cpp
// Access the PivotTable
intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));

// Access the data field
intrusive_ptr<PivotDataField> pivotDataField = pt->GetDataFields()->Get(0);

// Set data display format
pivotDataField->SetShowDataAs(PivotFieldFormatType::PercentageOfColumn);
```

---

# spire.xls cpp pivot table
## refresh pivot table data
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

//Update the data source of PivotTable.
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2"))->SetValue(L"999");

//Get the PivotTable that was built on the data source.
intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetPivotTables()->Get(0));

//Refresh the data of PivotTable.
pt->GetCache()->SetIsRefreshOnLoad(true);
```

---

# spire.xls cpp pivot table
## show data field in row for pivot table
```cpp
// Get the second worksheet from the workbook and retrieve the first pivot table
intrusive_ptr<XlsPivotTable>pivotTable = dynamic_pointer_cast<XlsPivotTable>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetPivotTables()->Get(0));

// Set the option to show the data field in the row for the pivot table
pivotTable->SetShowDataFieldInRow(true);

// Calculate the data for the pivot table
pivotTable->CalculateData();
```

---

# spire.xls cpp pivot table
## show subtotals in pivot table
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"Pivot Table"));

intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));

//Show Subtotals
pt->SetShowSubtotals(true);
```

---

# Update Pivot Table Data Source
## Update data source values and refresh pivot table
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> data = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"Data"));
intrusive_ptr<CellRange> a2 = dynamic_pointer_cast<CellRange>(data->Get(L"A2"));

a2->SetText(L"NewValue");
data->Get(L"D2")->SetNumberValue(28000);

//Get the sheet in which the pivot table is located
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"PivotTable"));

intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));
//Refresh and calculate
pt->GetCache()->SetIsRefreshOnLoad(true);
pt->CalculateData();
```

---

# spire.xls cpp tracked changes
## accept or reject tracked changes in excel
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Accept the changes or reject the changes.
//workbook.AcceptAllTrackedChanges();
workbook->RejectAllTrackedChanges();
```

---

# Detect Workbook Protection
## Check if an Excel workbook is password protected

```cpp
bool value = Workbook::IsPasswordProtected(inputFile.c_str());
wstring* boolvalue = new wstring();
if (value)
{
    boolvalue->append(L"Yes");
}
else
{
    boolvalue->append(L"No");
}
```

---

# Hide Formulas in Excel
## Hide formulas in worksheet and protect it with password
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Hide the formulas in the used range
sheet->GetAllocatedRange()->SetIsFormulaHidden(true);

//Protect the worksheet with password
sheet->XlsWorksheetBase::Protect(L"e-iceblue");
```

---

# excel cell locking
## lock specific cells in excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Create an empty worksheet.
workbook->CreateEmptySheet();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Loop through all the rows in the worksheet and unlock them.
for (int i = 0; i < 255; i++)
{
	sheet->GetRows()->GetItem(i)->GetStyle()->SetLocked(false);
}

//Lock specific cell in the worksheet.
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Locked");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->GetStyle()->SetLocked(true);

//Lock specific cell range in the worksheet.
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1:E3"))->SetText(L"Locked");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1:E3"))->GetStyle()->SetLocked(true);

//Set the password.
sheet->XlsWorksheetBase::Protect(L"123");
```

---

# excel sheet column locking
## lock specific column in excel worksheet
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Create an empty worksheet
workbook->CreateEmptySheet();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Loop through all the columns in the worksheet and unlock them
for (int i = 0; i < 255; i++)
{
	sheet->GetRows()->GetItem(i)->GetStyle()->SetLocked(false);
}

//Lock the fourth column in the worksheet
sheet->GetColumns()->GetItem(3)->SetText(L"Locked");
sheet->GetColumns()->GetItem(3)->GetStyle()->SetLocked(true);
//Set the password
sheet->XlsWorksheetBase::Protect(L"123");
```

---

# c++ excel row locking
## lock specific rows in excel worksheet with password protection
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Create an empty worksheet.
workbook->CreateEmptySheet();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Loop through all the rows in the worksheet and unlock them.
for (int i = 0; i < 255; i++)
{
	sheet->GetRows()->GetItem(i)->GetStyle()->SetLocked(false);
}

//Lock the third row in the worksheet.
sheet->GetRows()->GetItem(2)->SetText(L"Locked");
sheet->GetRows()->GetItem(2)->GetStyle()->SetLocked(true);

//Set the password.
sheet->XlsWorksheetBase::Protect(L"123");
```

---

# spire.xls cpp security
## Protect specific cells in Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Protect cell
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->GetStyle()->SetLocked(true);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->GetStyle()->SetLocked(false);

sheet->XlsWorksheetBase::Protect(L"TestPassword");
```

---

# Excel Worksheet Protection with Editable Ranges
## Demonstrates how to protect an Excel worksheet while allowing specific ranges to remain editable
```cpp
// Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

// Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

// Define the specified ranges to allow users to edit while sheet is protected
sheet->AddAllowEditRange(L"EditableRanges", dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4:E12")));

// Protect worksheet with a password.
sheet->XlsWorksheetBase::Protect(L"TestPassword");
```

---

# spire.xls cpp security
## protect workbook with password
```cpp
//Protect Workbook
workbook->Protect(L"e-iceblue");
```

---

# c++ unlock protect sheet
## Unprotect a worksheet in an Excel file
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Unlock the worksheet in a unlocked Excel file with null string.
sheet->XlsWorksheetBase::Unprotect(L"e-iceblue");
```

---

# Unlock Excel Worksheet
## Demonstrates how to unprotect a worksheet in an Excel file
```cpp
using namespace Spire::Xls;

int main() {
	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Assume an Excel file is loaded
	
	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Unlock the worksheet in an unlocked Excel file with null string.
	sheet->XlsWorksheetBase::Unprotect();

	//Assume the file will be saved
	workbook->Dispose();
}
```

---

# spire.xls cpp textbox
## extract text from a textbox in Excel
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the first textbox.
intrusive_ptr<XlsTextBoxShape> shape = dynamic_pointer_cast<XlsTextBoxShape>(sheet->GetTextBoxes()->Get(0));

//Extract text from the text box.
wstring* content = new wstring();
content->append(L"The text extracted from the TextBox is: \n");
content->append(shape->GetText());
```

---

# spire.xls cpp textbox
## get textbox by name
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Insert a TextBox
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetText(L"Name：");
intrusive_ptr<ITextBoxShape> textBox = sheet->GetTextBoxes()->AddTextBox(2, 2, 18, 65);

//Set the name 
textBox->SetName(L"FirstTextBox");

//Set string text for TextBox 
textBox->SetText(L"Spire.XLS for C++ is a professional Excel  C++ API that can be used to create, read, write and convert Excel files in any type of C++ application.\n");

//Get the TextBox by the name
intrusive_ptr<ITextBoxShape> FindTextBox = sheet->GetTextBoxes()->Get(L"FirstTextBox");

//Get the TextBox text 
wstring text = FindTextBox->GetText();

workbook->Dispose();
```

---

# spire.xls cpp textbox manipulation
## modify textbox text and alignment in Excel worksheet
```cpp
//Get the first textbox
intrusive_ptr<ITextBox> tb = sheet->GetTextBoxes()->Get(0);

//Change the text of textbox
tb->SetText(L"Spire.XLS for C++");

//Set the alignment of textbox as center
tb->SetHAlignment(CommentHAlignType::Center);
tb->SetVAlignment(CommentVAlignType::Center);
```

---

# spire.xls cpp textbox
## remove borderline of textbox
```cpp
//Create a new worksheet named "Remove Borderline" and add a chart to the worksheet.
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
sheet->SetName(L"Remove Borderline");
intrusive_ptr<Spire::Xls::Chart> chart = sheet->GetCharts()->Add();

//Create textbox1 in the chart and input text information.
intrusive_ptr<XlsTextBoxShape> textbox1 = dynamic_pointer_cast<XlsTextBoxShape>(chart->GetTextBoxes()->AddTextBox(50, 50, 100, 600));
textbox1->SetText(L"The solution with borderline");

//Create textbox2 in the chart, input text information and remove borderline.
intrusive_ptr<XlsTextBoxShape> textbox2 = dynamic_pointer_cast<XlsTextBoxShape>(chart->GetTextBoxes()->AddTextBox(1000, 50, 100, 600));
textbox2->SetText(L"The solution without borderline");
textbox2->GetLine()->SetWeight(0);
```

---

# spire.xls cpp textbox
## set font and background for textbox
```cpp
using namespace Spire::Xls;

//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the textbox which will be edited.
intrusive_ptr<XlsTextBoxShape> shape = dynamic_pointer_cast<XlsTextBoxShape>(sheet->GetTextBoxes()->Get(0));

//Set the font and background color for the textbox.
//Set font.
intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
//font.IsStrikethrough = true;
font->SetFontName(L"Century Gothic");
font->SetSize(10);
font->SetIsBold(true);
font->SetColor(Spire::Xls::Color::GetBlue());
intrusive_ptr<RichTextShape> tempVar = dynamic_pointer_cast<RichTextShape>(shape->GetRichText());
wstring text = shape->GetText();
tempVar->SetFont(0, text.size() - 1, font);

//Set background color
shape->GetFill()->SetFillType(ShapeFillType::SolidColor);
shape->GetFill()->SetForeKnownColor(ExcelColors::BlueGray);
```

---

# spire.xls cpp textbox
## set internal margin of textbox
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Add a textbox to the sheet and set its position and size.
intrusive_ptr<XlsTextBoxShape> textbox = dynamic_pointer_cast<XlsTextBoxShape>(sheet->GetTextBoxes()->AddTextBox(4, 2, 100, 300));

//Set the text on the textbox.
textbox->SetText(L"Insert TextBox in Excel and set the margin for the text");
textbox->SetHAlignment(CommentHAlignType::Center);
textbox->SetVAlignment(CommentVAlignType::Center);

//Set the inner margins of the contents.
textbox->SetInnerLeftMargin(1);
textbox->SetInnerRightMargin(3);
textbox->SetInnerTopMargin(1);
textbox->SetInnerBottomMargin(1);
```

---

# spire.xls cpp textbox
## set wrap text for textbox
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Get the text box
intrusive_ptr<XlsTextBoxShape> shape = dynamic_pointer_cast<XlsTextBoxShape>(sheet->GetTextBoxes()->Get(0));

//Set wrap text
shape->SetIsWrapText(true);
```

---

# spire.xls cpp worksheet
## activate worksheet in workbook
```cpp
//Get the second worksheet from the workbook
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

//Activate the sheet
sheet->Activate();
```

---

# spire xls cpp page breaks
## add horizontal page breaks in excel worksheet
```cpp
//Add page break in Excel file.
(dynamic_pointer_cast<HPageBreaksCollection>(sheet->GetHPageBreaks()))->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E4")));
(dynamic_pointer_cast<HPageBreaksCollection>(sheet->GetHPageBreaks()))->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4")));
```

---

# spire xls cpp worksheet
## add new worksheet to workbook
```cpp
//Add a new worksheet named AddedSheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Add(L"AddedSheet"));
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetText(L"This is a new sheet.");
```

---

# spire.xls cpp worksheet style
## apply style to worksheet
```cpp
//Create a cell style
intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");
style->SetColor(Spire::Xls::Color::GetLightBlue());
style->GetFont()->SetColor(Spire::Xls::Color::GetWhite());
style->GetFont()->SetSize(15);
style->GetFont()->SetIsBold(true);
//Apply the style to the first worksheet
sheet->ApplyStyle(style);
```

---

# Check Dialog Sheet in Excel
## Determine if a worksheet in an Excel file is a dialog sheet
```cpp
wstring* content = new wstring();

//Check if the worksheet is a dialog sheet.
if (sheet->GetType() == ExcelSheetType::DialogSheet)
{
    content->append(L"Worksheet is a Dialog Sheet!");
}
else
{
    content->append(L"Worksheet is not a Dialog Sheet!");
}
```

---

# spire.xls cpp worksheet copy
## copy worksheet from one workbook to another
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Define a pagesetup object based on the first worksheet.
intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());
//The first five rows are repeated in each page. It can be seen in print preview.
pageSetup->SetPrintTitleRows(L"$1:$5");
//Create another Workbook.
intrusive_ptr<Workbook> workbook1 = new Workbook();
//Get the first worksheet in the book.
intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook1->GetWorksheets()->Get(0));
//Copy worksheet to destination worsheet in another Excel file.
sheet1->CopyFrom(sheet);
```

---

# spire.xls cpp worksheet
## copy worksheet within workbook
```cpp
//Get the first and the second worksheets.
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Add(L"MySheet"));
intrusive_ptr<CellRange> sourceRange = dynamic_pointer_cast<CellRange>(sheet->GetAllocatedRange());

//Copy the first worksheet to the second one.
sheet->Copy(sourceRange, sheet1, sheet->GetFirstRow(), sheet->GetFirstColumn(), true);
```

---

# c++ copy visible worksheets
## copy only visible worksheets from one workbook to another
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Create a new workbook
intrusive_ptr<Workbook> workbookNew = new Workbook();
workbookNew->SetVersion(ExcelVersion::Version2013);
workbookNew->GetWorksheets()->Clear();

//Loop through the worksheets
for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
{
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
    //Judge if the worksheet is visible
    if (sheet->GetVisibility() == WorksheetVisibility::Visible)
    {
        //Copy the sheet to new workbook
        wstring name = sheet->GetName();
        workbookNew->GetWorksheets()->AddCopy(sheet);
    }
}
```

---

# Copy Worksheet in Excel
## Copy a worksheet from one workbook to another
```cpp
//Create a workbook
intrusive_ptr<Workbook> sourceWorkbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> srcWorksheet = dynamic_pointer_cast<Worksheet>(sourceWorkbook->GetWorksheets()->Get(0));

//Create a workbook
intrusive_ptr<Workbook> targetWorkbook = new Workbook();

//Add a new worksheet
intrusive_ptr<Worksheet> targetWorksheet = targetWorkbook->GetWorksheets()->Add(L"added");

//Copy the first worksheet of source Excel document to the new added worksheet of target Excel document
targetWorksheet->CopyFrom(srcWorksheet);

sourceWorkbook->Dispose();
```

---

# spire.xls cpp worksheet detection
## detect empty worksheets in excel workbook
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Get the first worksheet
intrusive_ptr<Worksheet> worksheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Detect the first worksheet is empty or not
bool detect1 = worksheet1->GetIsEmpty();

//Get the second worksheet
intrusive_ptr<Worksheet> worksheet2 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

//Detect the second worksheet is empty or not
bool detect2 = worksheet2->GetIsEmpty();
```

---

# spire.xls cpp worksheet
## fill data in worksheet with text and number values
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Fill data
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->GetStyle()->GetFont()->SetIsBold(true);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->GetFont()->SetIsBold(true);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->GetStyle()->GetFont()->SetIsBold(true);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Month");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetText(L"January");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetText(L"February");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetText(L"March");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetText(L"April");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetText(L"Payments");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(251);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(515);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(454);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(874);
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetText(L"Sample");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetText(L"Sample1");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetText(L"Sample2");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetText(L"Sample3");
dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetText(L"Sample4");

//Set width for the second column
sheet->SetColumnWidth(2, 10);
```

---

# spire.xls cpp worksheet
## freeze panes in excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Freeze Top Row
sheet->FreezePanes(2, 1);
```

---

# spire.xls cpp worksheet
## get freeze pane range information
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//The row and column index of the frozen pane is passed through the out parameter. 
//If it returns to 0, it means that it is not frozen
int rowIndex = sheet->GetFreezePanes()[0];
int colIndex = sheet->GetFreezePanes()[1];

wstring range = L"Row index: " + to_wstring(rowIndex) + L", column index: " + to_wstring(colIndex);
```

---

# spire.xls cpp fonts
## get list of fonts used in excel workbook
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

std::vector<intrusive_ptr<ExcelFont>> fonts;

//Loop all sheets of workbook
for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
{
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
    for (int r = 0; r < sheet->GetRows()->GetCount(); r++)
    {
        for (int c = 0; c < sheet->GetRows()->GetItem(r)->GetCells()->GetCount(); c++)
        {
            //Get the font of cell and add it to list
            fonts.push_back(dynamic_pointer_cast<ExcelFont>(sheet->GetRows()->GetItem(r)->GetCells()->GetItem(c)->GetStyle()->GetFont()));
        }
    }
}
wstring* strB = new wstring();

for (auto font : fonts)
{
    strB->append(L"FontName:");
    strB->append(font->GetFontName());
    strB->append(L"; FontSize:{1}");
    strB->append(to_wstring(font->GetSize()));
}
```

---

# get worksheet names
## retrieve names of all worksheets in an excel workbook
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(L"input_file_path.xlsx");

//Get the names of all worksheets
wstring* content = new wstring();
for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
{
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
    content->append(sheet->GetName());
}
```

---

# spire.xls cpp worksheet visibility
## Hide or show worksheet in Excel workbook
```cpp
//Hide the sheet named "Sheet1"
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"Sheet1"))->SetVisibility(WorksheetVisibility::Hidden);
//Show the second sheet
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->SetVisibility(WorksheetVisibility::Visible);
```

---

# spire.xls cpp worksheet
## hide worksheet tabs in Excel
```cpp
//Hide worksheet tab
workbook->SetShowTabs(false);
```

---

# spire.xls hide zero values
## hide zero values in Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set false to hide the zero values in sheet
sheet->SetIsDisplayZeros(false);
```

---

# spire.xls cpp document properties
## link custom document property to content
```cpp
//Add a custom document property
workbook->GetCustomDocumentProperties()->Add(L"Test", L"MyNamedRange");
//Get the added document property
intrusive_ptr<ICustomDocumentProperties> properties = workbook->GetCustomDocumentProperties();
intrusive_ptr<IDocumentProperty> property_Renamed = properties->Get(L"Test");
//Link to content 
property_Renamed->SetLinkToContent(true);
```

---

# spire.xls cpp worksheet
## move worksheet to a specific position
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Move worksheet
sheet->MoveWorksheet(2);
```

---

# c++ page break preview
## set zoom scale for page break view in excel
```cpp
//Set the scale of PageBreakView mode in Excel file.
sheet->SetZoomScalePageBreakView(80);
```

---

# Spire.XLS C++ Page Break Management
## Remove vertical and horizontal page breaks in an Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Clear all the vertical page breaks
dynamic_pointer_cast<VPageBreaksCollection>(sheet->GetVPageBreaks())->Clear();

//Remove the first horizontal Page Break
dynamic_pointer_cast<HPageBreaksCollection>(sheet->GetHPageBreaks())->RemoveAt(0);

//Set the ViewMode as Preview to see how the page breaks work
sheet->SetViewMode(ViewMode::Preview);
```

---

# spire.xls cpp worksheet
## remove worksheet from excel workbook
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

// Remove a worksheet by sheet index
workbook->GetWorksheets()->RemoveAt(1);
```

---

# spire.xls cpp page break
## set horizontal page breaks and view mode
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set Excel Page Break Horizontally
(dynamic_pointer_cast<HPageBreaksCollection>(sheet->GetHPageBreaks()))->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8")));
(dynamic_pointer_cast<HPageBreaksCollection>(sheet->GetHPageBreaks()))->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A14")));

//Set view mode to Preview mode
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->SetViewMode(ViewMode::Preview);
```

---

# spire.xls worksheet tab color
## Set worksheet tab colors in Excel
```cpp
//Set the tab color of first sheet to be red 
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
sheet->SetTabColor(Spire::Xls::Color::GetRed());

//Set the tab color of first sheet to be green 
sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));
sheet->SetTabColor(Spire::Xls::Color::GetGreen());

//Set the tab color of first sheet to be blue 
sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(2));
sheet->SetTabColor(Spire::Xls::Color::GetLightBlue());
```

---

# spire.xls cpp worksheet view mode
## set worksheet view mode to preview
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Set the view mode 
dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->SetViewMode(ViewMode::Preview);
```

---

# spire.xls cpp worksheet gridlines
## Show or hide gridlines in Excel worksheets
```cpp
//Get the first and second worksheet
intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
intrusive_ptr<Worksheet> sheet2 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

//Hide grid line in the first worksheet
sheet1->SetGridLinesVisible(false);
//Show grid line in the first worksheet
sheet2->SetGridLinesVisible(true);
```

---

# spire.xls cpp worksheet
## show worksheet tabs in excel workbook
```cpp
//Show worksheet tab
workbook->SetShowTabs(true);
```

---

# Split Worksheet Into Panes
## This code demonstrates how to split an Excel worksheet into multiple panes, both vertically and horizontally, and set the active pane.
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Vertical and horizontal split the worksheet into four panes
sheet->SetFirstVisibleColumn(2);
sheet->SetFirstVisibleRow(5);
sheet->SetVerticalSplit(4000);
sheet->SetHorizontalSplit(5000);

//Set the active pane
sheet->SetActivePane(1);
```

---

# Unfreeze Excel Panes
## Demonstrates how to unfreeze panes in an Excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Unfreeze the panes.
sheet->RemovePanes();
```

---

# spire.xls cpp worksheet
## verify if worksheet is password protected
```cpp
//Get the first worksheet from workbook
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Verify the first worksheet 
bool detect = sheet->GetIsPasswordProtected();
```

---

# spire.xls cpp worksheet zoom
## set worksheet zoom factor
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set the zoom factor of the sheet to 85
sheet->SetZoom(85);
```

---

# spire.xls cpp custom properties
## add custom properties to Excel workbook
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Add a custom property to make the document as final
workbook->GetCustomDocumentProperties()->Add(L"_MarkAsFinal", true);

//Add other custom properties to the workbook
workbook->GetCustomDocumentProperties()->Add(L"The Editor", L"E-iceblue");
workbook->GetCustomDocumentProperties()->Add(L"Phone number", 81705109);
workbook->GetCustomDocumentProperties()->Add(L"Revision number", 7.12);
tm t;
t.tm_year = 2021 - 1900;
t.tm_mon = 1 - 1;
t.tm_mday = 8;
t.tm_hour = 8;//beijing zone must +8
t.tm_min = 0;
t.tm_sec = 0;
intrusive_ptr<Spire::Xls::DateTime> dt = new Spire::Xls::DateTime(2021, 1, 8, 0, 0, 0);
workbook->GetCustomDocumentProperties()->Add(L"Revision date", dt);
```

---

# spire.xls cpp decrypt workbook
## decrypt password protected Excel workbook
```cpp
// Check if the file is password protected
bool value = Workbook::IsPasswordProtected(L"DecryptWorkbook.xlsx");

if (value)
{
    // Load a file with the password specified
    intrusive_ptr<Workbook> workbook = new Workbook();
    workbook->SetOpenPassword(L"eiceblue");
    workbook->LoadFromFile(L"DecryptWorkbook.xlsx");

    // Decrypt workbook
    workbook->UnProtect();
    
    workbook->Dispose();
}
```

---

# Detect VBA Macros in Excel Workbook
## Check if an Excel file contains VBA macros using Spire.XLS for C++
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Detect if the Excel file contains VBA macros
wstring value = L"";
bool hasMacros = workbook->GetHasMacros();
if (hasMacros)
{
    value = L"Yes";
}
else
{
    value = L"No";
}

workbook->Dispose();
```

---

# Spire.XLS C++ Encrypt Workbook
## How to protect and encrypt an Excel workbook with password
```cpp
using namespace Spire::Xls;

int main() {
    //Create a workbook
    intrusive_ptr<Workbook> workbook = new Workbook();

    //Protect Workbook with the password you want
    workbook->Protect(L"eiceblue");

    workbook->Dispose();
}
```

---

# Hide Excel Window
## Hide Excel window using Spire.XLS for C++
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(L"example.xlsx");

//Hide window
workbook->SetIsHideWindow(true);
```

---

# Load and Save Excel File with Macro
## This example demonstrates how to load an Excel file with macro, modify its content, and save it back to a new file.
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(L"MacroSample.xls");

//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetText(L"This is a simple test!");

//Save to file.
workbook->SaveToFile(L"LoadAndSaveFileWithMacro.xls", ExcelVersion::Version97to2003);
workbook->Dispose();
```

---

# spire.xls cpp merge excel files
## merge multiple excel files into one workbook
```cpp
//Create a workbook
intrusive_ptr<Workbook> newbook = new Workbook();
newbook->SetVersion(ExcelVersion::Version2013);
//Clear all worksheets
newbook->GetWorksheets()->Clear();

//Create a workbook
intrusive_ptr<Workbook> tempbook = new Workbook();

for (auto file : files)
{
    //Load the file
    tempbook->LoadFromFile(file.c_str());
    for (int i = 0; i < tempbook->GetWorksheets()->GetCount(); i++)
    {
        intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(tempbook->GetWorksheets()->Get(i));
        //Copy every sheet in a workbook
        (dynamic_pointer_cast<XlsWorksheetsCollection>(newbook->GetWorksheets()))->AddCopy(sheet, WorksheetCopyType::CopyAll);
    }
}

//Save to file.
workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
workbook->Dispose();
```

---

# spire.xls cpp encrypted file
## open encrypted Excel file with password attempts
```cpp
using namespace Spire::Xls;

std::wstring inputFile = L"EncryptedFile.xlsx";
vector<wstring> passwords = { L"password1",  L"password2",  L"password3",  L"1234" };
for (int i = 0; i < passwords.size(); i++)
{
    try
    {
        //Create a workbook
        intrusive_ptr<Workbook> workbook = new Workbook();

        //Set open password
        workbook->SetOpenPassword(passwords[i].c_str());

        //Load the encrypted document
        workbook->LoadFromFile(inputFile.c_str());
    }
    catch (SpireException ex)
    {
        // Password is incorrect
    }
}
```

---

# spire.xls cpp file operations
## demonstrate different ways to open Excel files
```cpp
//1. Load file by file path
//Create a workbook
intrusive_ptr<Workbook> workbook1 = new Workbook();
//Load the document from disk
workbook1->LoadFromFile(inputFile.c_str());

//2. Load file by file stream 
ifstream inputf(inputFile.c_str(), ios::in | ios::binary);
intrusive_ptr<Stream> stream = new Stream(inputf);
//Create a workbook
intrusive_ptr<Workbook> workbook2 = new Workbook();
//Load the document from disk
workbook2->LoadFromStream(stream);

//3. Open Microsoft Excel 97 - 2003 file
intrusive_ptr<Workbook> wbExcel97 = new Workbook();
wbExcel97->LoadFromFile(inputFile_97.c_str(), ExcelVersion::Version97to2003);

//4. Open xml file
intrusive_ptr<Workbook> wbXML = new Workbook();
wbXML->LoadFromXml(inputFile_xml.c_str());

//5. Open csv file
intrusive_ptr<Workbook> wbCSV = new Workbook();
wbCSV->LoadFromFile(inputFile_csv.c_str(), L",", 1, 1);
```

---

# c++ excel stream reading
## load excel workbook from stream
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Open excel from a stream
ifstream inputf(inputFile.c_str(), ios::in | ios::binary);
intrusive_ptr<Stream> stream = new Stream(inputf);

workbook->LoadFromStream(stream);
```

---

# spire.xls cpp workbook
## remove custom properties from Excel workbook
```cpp
//Retrieve a list of all custom document properties of the Excel file
intrusive_ptr<ICustomDocumentProperties> customDocumentProperties = workbook->GetCustomDocumentProperties();

//Remove "Editor" custom document property
customDocumentProperties->Remove(L"Editor");
```

---

# spire.xls cpp save files
## save excel workbook to xlsx format
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Load the Excel document from disk
workbook->LoadFromFile(inputFile.c_str());

//Save to file.
workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
workbook->Dispose();
```

---

# spire.xls cpp stream
## save workbook to stream
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Save an excel workbook to stream
intrusive_ptr<Spire::Xls::Stream> stream = new Spire::Xls::Stream();
workbook->SaveToStream(stream);
```

---

# spire.xls cpp calculation mode
## Set Excel calculation mode to manual
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();

//Set excel calculation mode as Manual
workbook->SetCalculationMode(ExcelCalculationMode::Manual);
```

---

# spire.xls cpp worksheet margins
## set page margins for excel worksheet
```cpp
//Get the first worksheet
intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

//Set margins for top, bottom, left and right, here the unit of measure is Inch
sheet->GetPageSetup()->SetTopMargin(0.3);
sheet->GetPageSetup()->SetBottomMargin(1);
sheet->GetPageSetup()->SetLeftMargin(0.2);
sheet->GetPageSetup()->SetRightMargin(1);
//Set the header margin and footer margin
sheet->GetPageSetup()->SetHeaderMarginInch(0.1);
sheet->GetPageSetup()->SetFooterMarginInch(0.5);
```

---

# c++ excel theme setting
## set theme color in excel workbook
```cpp
//Create a workbook
intrusive_ptr<Workbook> workbook = new Workbook();
workbook->GetWorksheets()->Clear();
workbook->GetWorksheets()->AddCopy(srcWorksheet);

//1. Copy the theme of the workbook
//workbook->CopyTheme(srcWorkbook);

//2. Set a certain type of color of the default theme in the workbook
workbook->SetThemeColor(ThemeColorType::Dk1, Spire::Xls::Color::GetSkyBlue());
```

---

# spire.xls cpp track changes
## accept or reject tracked changes in excel workbook
```cpp
//Accept the changes or reject the changes.
//workbook.AcceptAllTrackedChanges();
workbook->RejectAllTrackedChanges();
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


