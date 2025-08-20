using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet;

namespace UsingStyleObject
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
        
            // Add a new worksheet to the Excel object
            Worksheet sheet = workbook.Worksheets.Add("new sheet");

            // Access the "B1" cell from the worksheet
            CellRange cell = sheet.Range["B1"];
      
            // Add some value to the "B1" cell
            cell.Text = "Hello Spire!";

            // Create a new style
            CellStyle style = workbook.Styles.Add("newStyle");

            // Set the vertical alignment of the text in the "B1" cell
            style.VerticalAlignment = VerticalAlignType.Center;

            // Set the horizontal alignment of the text in the "B1" cell
            style.HorizontalAlignment = HorizontalAlignType.Center;

            // Set the font color of the text in the "B1" cell
            style.Font.Color = Color.Blue;

            // Shrink the text to fit in the cell
            style.ShrinkToFit = true;

            // Set the bottom border color of the cell to GreenYellow
            style.Borders[BordersLineType.EdgeBottom].Color = Color.GreenYellow;

            // Set the bottom border type of the cell to Medium
            style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Medium;

            // Assign the Style object to the "B1" cell
            cell.Style = style;
         
            // Apply the same style to some other cells
            sheet.Range["B4"].Style = style;
            sheet.Range["B4"].Text = "Test";
            sheet.Range["C3"].CellStyleName = style.Name;
            sheet.Range["C3"].Text = "Welcome to use Spire.XLS";
            sheet.Range["D4"].Style = style;

            // Specify the name for the resulting Excel file
            String result = "UsingStyleObject_result.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
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
    }
}
