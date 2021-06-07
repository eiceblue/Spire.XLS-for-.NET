using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace CloneExcelFontStyle
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a workbook.
			Workbook workbook = new Workbook();

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Add the text to the Excel sheet cell range A1.
            sheet.Range["A1"].Text = "Text1";

            //Set A1 cell range's CellStyle.
            CellStyle style = workbook.Styles.Add("style");
            style.Font.FontName = "Calibri";
            style.Font.Color = Color.Red;
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.IsItalic = true;
            sheet.Range["A1"].CellStyleName = style.Name;

            //Clone the same style for B2 cell range.
            CellStyle csOrieign = style.clone();
            sheet.Range["B2"].Text = "Text2";
            sheet.Range["B2"].CellStyleName = csOrieign.Name;

            //Clone the same style for C3 cell range and then reset the font color for the text.
            CellStyle csGreen = style.clone();
            csGreen.Font.Color = Color.Green;
            sheet.Range["C3"].Text = "Text3";
            sheet.Range["C3"].CellStyleName = csGreen.Name;

            String result = "Result-CloneExcelFontStyle.xlsx";

            //Save to file.
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            //Launch the MS Excel file.
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
