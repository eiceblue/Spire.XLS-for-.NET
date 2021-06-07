using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateBubbleChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateBubbleChart.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            Chart chart = sheet.Charts.Add(ExcelChartType.Bubble);

            //Chart title
            chart.ChartTitle = "Bubble";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;

            chart.DataRange = sheet.Range["A1:C5"];
            chart.SeriesDataFromRange = false;
           
            chart.Series[0].Bubbles = sheet.Range["C2:C5"];
           
            //Set position of chart
            chart.LeftColumn = 7;
            chart.TopRow = 6;
            chart.RightColumn = 16;
            chart.BottomRow = 29;

            workbook.SaveToFile("CreateBubbleChart.xlsx", ExcelVersion.Version2010);

            //View the document
            FileViewer("CreateBubbleChart.xlsx");
        }

        private void FileViewer(string fileName)
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
