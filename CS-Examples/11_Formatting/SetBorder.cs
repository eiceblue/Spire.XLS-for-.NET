using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SetBorder
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

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetBorder.xlsx");

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

            // Specify the name for the resulting Excel file
            string result = "SetBorder_result.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result,ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
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
