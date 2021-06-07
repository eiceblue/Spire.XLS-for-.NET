using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace FitWidthWhenConvertToPDF
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

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                //Auto fit page height
                sheet.PageSetup.FitToPagesTall = 0;
                //Fit one page width
                sheet.PageSetup.FitToPagesWide = 1;
            }

            //Save and launch result file
            string result = "result.pdf";
            workbook.SaveToFile(result, FileFormat.PDF);
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
