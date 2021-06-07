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
            //Load a Workbook from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ImageHeaderFooter.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Load an image from disk
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");

            //Set the image header
            sheet.PageSetup.LeftHeaderImage = image;
            sheet.PageSetup.LeftHeader = "&G";

            //Set the image footer
            sheet.PageSetup.CenterFooterImage = image;
            sheet.PageSetup.CenterFooter = "&G";

            //Set the view mode of the sheet
            sheet.ViewMode = ViewMode.Layout;

            String result="Output_ImageHeaderFooter.xlsx";
            //Save and Launch
            workbook.SaveToFile(result, ExcelVersion.Version2010);
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
