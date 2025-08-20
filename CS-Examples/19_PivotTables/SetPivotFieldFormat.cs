using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace SetPivotFieldFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load an excel file including pivot table
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PivotTableExample.xlsx");

            // Get the sheet in which the pivot table is located
            Worksheet sheet = workbook.Worksheets["PivotTable"];
           
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

            // Specify the filename for the resulting workbook
            string result = "SetPivotFieldFormat_result.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
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
