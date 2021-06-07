using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace FormatAxis
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a Workbook
            Workbook workbook = new Workbook();

            //Get the first sheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "FormatAxis";

            //Set chart data
            CreateChartData(sheet);

            //Add a chart
            Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered);
            chart.DataRange = sheet.Range["B1:B9"];
            chart.SeriesDataFromRange = false;
            chart.PlotArea.Visible = false;
            chart.TopRow = 10;
            chart.BottomRow = 28;
            chart.LeftColumn = 2;
            chart.RightColumn = 10;
            chart.ChartTitle = "Chart with Customized Axis";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;
            Spire.Xls.Charts.ChartSerie cs1 = chart.Series[0];
            cs1.CategoryLabels = sheet.Range["A2:A9"];

            //Format axis
            chart.PrimaryValueAxis.MajorUnit = 8;
            chart.PrimaryValueAxis.MinorUnit = 2;
            chart.PrimaryValueAxis.MaxValue = 50;
            chart.PrimaryValueAxis.MinValue = 0;
            chart.PrimaryValueAxis.IsReverseOrder = false;
            chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkOutside;
            chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkInside;
            chart.PrimaryValueAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionNextToAxis;
            chart.PrimaryValueAxis.CrossesAt = 0;

            //Set NumberFormat
            chart.PrimaryValueAxis.NumberFormat = "$#,##0";
            chart.PrimaryValueAxis.IsSourceLinked = false;

            ChartSerie serie = chart.Series[0];

            foreach (ChartDataPoint dataPoint in serie.DataPoints)
            {
                //Format Series
                dataPoint.DataFormat.Fill.FillType = ShapeFillType.SolidColor;
                dataPoint.DataFormat.Fill.ForeColor = Color.LightGreen;

                //Set transparency
                dataPoint.DataFormat.Fill.Transparency =0.3;           
            }
            
            //Save and Launch
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010);
            ExcelDocViewer("Output.xlsx");
        }
        private void CreateChartData(Worksheet sheet)
        {
            //Set value of specified cell
            sheet.Range["A1"].Value = "Month";
            sheet.Range["A2"].Value = "Jan";
            sheet.Range["A3"].Value = "Feb";
            sheet.Range["A4"].Value = "Mar";
            sheet.Range["A5"].Value = "Apr";
            sheet.Range["A6"].Value = "May";
            sheet.Range["A7"].Value = "Jun";
            sheet.Range["A8"].Value = "Jul";
            sheet.Range["A9"].Value = "Aug";

            sheet.Range["B1"].Value = "Planned";
            sheet.Range["B2"].NumberValue = 38;
            sheet.Range["B3"].NumberValue = 47;
            sheet.Range["B4"].NumberValue = 39;
            sheet.Range["B5"].NumberValue = 36;
            sheet.Range["B6"].NumberValue = 27;
            sheet.Range["B7"].NumberValue = 25;
            sheet.Range["B8"].NumberValue = 36;
            sheet.Range["B9"].NumberValue = 48;
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
