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
            //Create a workbook
            Workbook workbook = new Workbook();

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Insert a picture
            ExcelPicture pic = sheet.Pictures.Add(2, 2,@"..\..\..\..\..\..\Data\logo.png");

            //Set left offset and top offset from the current range
            pic.LeftColumnOffset = 200;
            pic.TopRowOffset = 100;

            //Save the Excel file
            string result = "PictureOffset_out.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

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
