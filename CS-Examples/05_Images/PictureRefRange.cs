using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace PictureRefRange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PictureRefRange.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Range["A1"].Value = "Spire.XLS";
            sheet.Range["B3"].Value = "E-iceblue";

            //Get the first picture in worksheet
            ExcelPicture picture = sheet.Pictures[0];

            //Set the reference range of the picture to A1:B3
            picture.RefRange = "A1:B3";

            //Save the Excel file
            string result = "PictureRefRange_out.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            //Launch the Excel file
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
