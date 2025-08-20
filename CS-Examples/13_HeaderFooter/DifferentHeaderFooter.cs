using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DifferentHeaderFooter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a Workbook
            Workbook wb = new Workbook();

            // Load file from disk
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\DifferentHeaderFooter.xlsx");

            // Get the first worksheet
            Worksheet sheet = wb.Worksheets[0];

            // Set text for the range
            sheet.Range["A1"].Text = "Page 1";
            sheet.Range["G1"].Text = "Page 2";

            // Set the different header footer for Odd and Even pages
            sheet.PageSetup.DifferentOddEven = 1;

            // Set the header with font, size, bold and color
            sheet.PageSetup.OddHeaderString = "&\"Arial\"&12&B&KFFC000 Odd_Header";
            sheet.PageSetup.OddFooterString = "&\"Arial\"&12&B&KFFC000 Odd_Footer";
            sheet.PageSetup.EvenHeaderString = "&\"Arial\"&12&B&KFF0000 Even_Header";
            sheet.PageSetup.EvenFooterString = "&\"Arial\"&12&B&KFF0000 Even_Footer";

            // Set view mode as  page layout view
            sheet.ViewMode = ViewMode.Layout;

            // Save the workbook to the specified file in Excel 2013 format
            wb.SaveToFile("Output.xlsx", ExcelVersion.Version2013);

            // Launch the file
            ExcelDocViewer("Output.xlsx");
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
