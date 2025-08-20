using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace PictureOffset
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

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a picture
            ExcelPicture pic = sheet.Pictures.Add(2, 2,@"..\..\..\..\..\..\Data\logo.png");

            // Set the left offset and top offset of the picture from the current range.
            pic.LeftColumnOffset = 200;
            pic.TopRowOffset = 100;

            // Save the modified workbook to a file named "PictureOffset_out.xlsx" using Excel 2013 format.
            string result = "PictureOffset_out.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources.
            workbook.Dispose();

            //Launch the file
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
