using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace ToImageWithoutWhiteSpace
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Load the document from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set the margin as 0 to remove the white space around the image
            sheet.PageSetup.LeftMargin = 0;
            sheet.PageSetup.BottomMargin = 0;
            sheet.PageSetup.TopMargin = 0;
            sheet.PageSetup.RightMargin = 0;

            //convert to image
            Image image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn);

            //Save and launch result file
            string result = "result.png";
            image.Save(result);
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
