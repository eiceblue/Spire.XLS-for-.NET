using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace ApplyBuiltInStyles
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {

            Workbook workbook = new Workbook();
            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Apply title style
            sheet.Range["A1:J1"].BuiltInStyle = BuiltInStyles.Title;

            //Save and launch result file
            string result = "result.xlsx";
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
