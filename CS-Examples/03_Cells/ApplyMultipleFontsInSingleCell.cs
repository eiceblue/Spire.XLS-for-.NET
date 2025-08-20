using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace ApplyMultipleFontsInSingleCell
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

            //Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_1.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Create a font object in workbook, setting the font color, size and type.
            ExcelFont font1 = workbook.CreateFont();
            font1.KnownColor = ExcelColors.LightBlue;
            font1.IsBold = true;
            font1.Size = 10;

            //Create another font object specifying its properties.
            ExcelFont font2 = workbook.CreateFont();
            font2.KnownColor = ExcelColors.Red;
            font2.IsBold = true;
            font2.IsItalic = true;
            font2.FontName = "Times New Roman";
            font2.Size = 11;

            //Write a RichText string to the cell 'A1', and set the font for it.
            RichText richText = sheet.Range["H5"].RichText;
            richText.Text = "This document was created with Spire.XLS for .NET.";
            richText.SetFont(0, 29, font1);
            richText.SetFont(31, 48, font2);

            //Specify the filename for the resulting Excel file
            String result = "Result-ApplyMultipleFontsInSingleCell.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013); ;

            // Dispose of the workbook object to release resources
            workbook.Dispose();

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
