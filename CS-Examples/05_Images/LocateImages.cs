using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace LocateImages
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a Workbook
            Workbook workbook = new Workbook();

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\LocateImages.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first picture from the sheet.
            ExcelPicture pic = sheet.Pictures[0];

            // Set the horizontal offset of the picture within the cell to 300.
            pic.LeftColumnOffset = 300;

            // Set the vertical offset of the picture within the cell to 300.
            pic.TopRowOffset = 300;

            // Save the modified workbook to a file named "Output.xlsx" using Excel 2010 format.
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources.
            workbook.Dispose();

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
