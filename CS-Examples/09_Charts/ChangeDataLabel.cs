using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace ChangeDataLabel
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeDataLabel.xlsx");

            Worksheet sheet = workbook.Worksheets[0];

            //Get the chart
            Chart chart = sheet.Charts[0];

            //Change data label of the frist datapoint of the first series
            chart.Series[0].DataPoints[0].DataLabels.Text = "changed data label";

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
