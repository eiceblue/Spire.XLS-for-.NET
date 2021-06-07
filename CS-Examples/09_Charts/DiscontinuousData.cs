using Spire.Xls;
using Spire.Xls.Charts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DiscontinuousData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a Workbook from disk
            Workbook book = new Workbook();
            book.LoadFromFile(@"..\..\..\..\..\..\Data\DiscontinuousData.xlsx");

            //Get the first sheet
            Worksheet sheet = book.Worksheets[0];

            //Add a chart
            Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered);
            chart.SeriesDataFromRange = false;

            //Set the position of chart
            chart.LeftColumn = 1;
            chart.TopRow = 10;
            chart.RightColumn = 10;
            chart.BottomRow = 24;

            //Add a series
            ChartSerie cs1 = (ChartSerie)chart.Series.Add();

            //Set the name of the cs1
            cs1.Name = sheet.Range["B1"].Value;

            //Set discontinuous values for cs1
            cs1.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"]);
            cs1.Values = sheet.Range["B2:B3"].AddCombinedRange(sheet.Range["B5:B6"]).AddCombinedRange(sheet.Range["B8:B9"]);

            //Set the chart type
            cs1.SerieType = ExcelChartType.ColumnClustered;

            //Add a series
            ChartSerie cs2 = (ChartSerie)chart.Series.Add();
            cs2.Name = sheet.Range["C1"].Value;
            cs2.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"]);
            cs2.Values = sheet.Range["C2:C3"].AddCombinedRange(sheet.Range["C5:C6"]).AddCombinedRange(sheet.Range["C8:C9"]);
            cs2.SerieType = ExcelChartType.ColumnClustered;

            chart.ChartTitle = "Chart";
            chart.ChartTitleArea.Font.Size = 20;
            chart.ChartTitleArea.Color = Color.Black;

            chart.PrimaryValueAxis.HasMajorGridLines = false;

            //Save and Launch
            book.SaveToFile("Output.xlsx",ExcelVersion.Version2010);
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
