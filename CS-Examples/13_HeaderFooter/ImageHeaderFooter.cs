using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ImageHeaderFooter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load a Workbook from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ImageHeaderFooter.xlsx");

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Load an image from disk
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");

			//////////////////Use the following code for netstandard dlls/////////////////////////
			/*
			SkiaSharp.SKBitmap image = SkiaSharp.SKBitmap.Decode((@"..\..\..\..\..\..\Data\Logo.png");
			*/
            
            // Set the image header
            sheet.PageSetup.LeftHeaderImage = image;
            sheet.PageSetup.LeftHeader = "&G";

            // Set the image footer
            sheet.PageSetup.CenterFooterImage = image;
            sheet.PageSetup.CenterFooter = "&G";

            // Set the view mode of the sheet
            sheet.ViewMode = ViewMode.Layout;

            // Specify the file name for the resulting Excel file
            String result ="Output_ImageHeaderFooter.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
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
