using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace ShowLeaderLine
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
            workbook.Version = ExcelVersion.Version2013;

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set value of specified range
            sheet.Range["A1"].Value = "1";
            sheet.Range["A2"].Value = "2";
            sheet.Range["A3"].Value = "3";
            sheet.Range["B1"].Value = "4";
            sheet.Range["B2"].Value = "5";
            sheet.Range["B3"].Value = "6";
            sheet.Range["C1"].Value = "7";
            sheet.Range["C2"].Value = "8";
            sheet.Range["C3"].Value = "9";

            Chart chart = sheet.Charts.Add(ExcelChartType.BarStacked);
            chart.DataRange = sheet.Range["A1:C3"];
            chart.TopRow = 4;
            chart.LeftColumn = 2;
            chart.Width = 450;
            chart.Height = 300;

            foreach (ChartSerie cs in chart.Series)
            {
                cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
                cs.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = true;
            }

            workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2013);
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
