using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ApplyStyleForUsedCell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new Workbook object.
            Workbook workbook = new Workbook();

            // Load an existing Excel file into the workbook.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

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

            // Define the filename for the resulting Excel file.
            string result = "ApplyStyle_result.xlsx";

            // Save the modified workbook to a new file with the specified filename and Excel version.
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // View the document
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
