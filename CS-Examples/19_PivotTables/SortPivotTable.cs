using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SortPivotTable
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

            // Load an Excel file that contains a pivot table
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SortPivotTable.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

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

            // Specify the output file name for saving the modified workbook
            String result = "SortPivotTable_result.xlsx";

            // Save the workbook to the specified file, using the Excel 2013 file format
            workbook.SaveToFile(result, ExcelVersion.Version2013);

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
