using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;
using System.Text;
using System.IO;

namespace AddImageToFirstPageHeaderFooter
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddImageToFirstPageHeaderFooter.xlsx");

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.DifferentFirst = (byte)1;

            // Load an image from disk
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");

            // Set the image header
            sheet.PageSetup.FirstLeftHeaderImage = image;
            sheet.PageSetup.FirstCenterHeaderImage = image;
            sheet.PageSetup.FirstRightHeaderImage = image;

            // Set the image footer
            sheet.PageSetup.FirstLeftFooterImage = image;
            sheet.PageSetup.FirstCenterFooterImage = image;
            sheet.PageSetup.FirstRightFooterImage = image;

            // Set the view mode of the sheet
            sheet.ViewMode = ViewMode.Layout;

            // Specify the file name for the resulting Excel file
            String result = "Output_AddImageHeaderFooterToFirstPage.xlsx";

            // Save the workbook to the specified file in Excel 2016 format
            workbook.SaveToFile(result, ExcelVersion.Version2016);

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
