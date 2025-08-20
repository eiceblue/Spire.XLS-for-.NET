using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet;

namespace FormatARow
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
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

            String result = "FormatARow_result.xlsx";

            //Save to file
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
