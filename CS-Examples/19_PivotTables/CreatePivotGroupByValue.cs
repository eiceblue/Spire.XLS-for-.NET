using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using System;
using System.Windows.Forms;

namespace CreatePivotGroupByValue
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            workbook.LoadFromFile("..\\..\\..\\..\\..\\..\\Data\\CreatePivotGroupByValue.xlsx");

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

            // Specify the filename for the resulting Excel file
            String result = "CreatePivotGroupByValue-out.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2016);

            // Dispose of the workbook object
            workbook.Dispose();

            // View the document using a file viewer
            FileViewer(result);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
