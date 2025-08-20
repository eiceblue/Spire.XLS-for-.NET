using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace RepeatItemLabels
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\RepeatItemLabelsExample.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add an empty worksheet 
            Worksheet sheet2 = workbook.CreateEmptySheet();
            sheet2.Name = "Pivot Table";

            // Define the data range for the pivot table
            CellRange dataRange = sheet.Range["A1:D9"];

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

            // Specify the filename for the resulting workbook
            String result = "RepeatItemLabels_result.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the document
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
