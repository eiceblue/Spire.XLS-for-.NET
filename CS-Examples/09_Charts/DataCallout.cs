using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace DataCallout
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

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataCallout.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the first chart
            Chart chart = sheet.Charts[0];

            // Enable data labels and customize callout settings for each series in the chart
            foreach (ChartSerie cs in chart.Series)
            {
                cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
                cs.DataPoints.DefaultDataPoint.DataLabels.HasWedgeCallout = true;
                cs.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = true;
                cs.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = true;
                cs.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = true;
            }

            //Save the file
            workbook.SaveToFile("Output.xlsx", FileFormat.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(workbook.FileName);
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
