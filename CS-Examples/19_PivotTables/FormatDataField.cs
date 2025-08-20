using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace FormatDataField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FormatDataField.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first pivot table from the sheet
            XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;
            // Access the data field.
            PivotDataField pivotDataField = pt.DataFields[0];

            // Set data display format
            pivotDataField.ShowDataAs = PivotFieldFormatType.PercentageOfColumn;

            // Specify the filename for the resulting workbook
            String result = "FormatDataField_output.xlsx";

            // Save the modified workbook to a file using Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Launch the file
            ExcelDocViewer(result);
        }

        private void ExcelDocViewer(string fileName)
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
