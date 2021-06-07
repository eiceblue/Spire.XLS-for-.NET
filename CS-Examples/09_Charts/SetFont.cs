using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace SetFont
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Load a Workbook from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetFont.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the first sheet
            Chart chart = sheet.Charts[0];

            //Create a font
            ExcelFont font = workbook.CreateFont();
            font.Size = 15.0;
            font.Color = Color.LightSeaGreen;

            foreach (ChartSerie cs in chart.Series)
            {
                //Set font
                cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font);
            }

            //Save and Launch
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010);
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
