using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;


namespace ChangeDataRange
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampeB_4.xlsx");

            Worksheet sheet = workbook.Worksheets[0];

            //Get chart
            Chart chart = sheet.Charts[0];

            //Change data range
            chart.DataRange = sheet.Range["A1:C4"];

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
