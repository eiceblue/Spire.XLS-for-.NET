using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet;

namespace ForegroundAndBackground
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
            CellStyle style = workbook.Styles.Add("newStyle1");

            //Set filling pattern type
            style.Interior.FillPattern = ExcelPatternType.VerticalStripe;

            //Set filling Background color
            style.Interior.Gradient.BackKnownColor = ExcelColors.Green;

            //Set filling Foreground color
            style.Interior.Gradient.ForeKnownColor = ExcelColors.Yellow;

            //Apply the style to  "B2" cell
            sheet.Range["B2"].CellStyleName = style.Name;
            sheet.Range["B2"].Text = "Test";
            sheet.Range["B2"].RowHeight = 30;
            sheet.Range["B2"].ColumnWidth = 50;


            //Create a new style
            style = workbook.Styles.Add("newStyle2");

            //Set filling pattern type
            style.Interior.FillPattern = ExcelPatternType.ThinHorizontalStripe;

            //Set filling Foreground color
            style.Interior.Gradient.ForeKnownColor = ExcelColors.Red;

            //Apply the style to  "B4" cell
            sheet.Range["B4"].CellStyleName = style.Name;
            sheet.Range["B4"].RowHeight = 30;
            sheet.Range["B4"].ColumnWidth = 60;

            String result = "ForegroundAndBackground_result.xlsx";

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
