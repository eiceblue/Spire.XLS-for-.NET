using Spire.Xls;
using Spire.Xls.Charts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ChangeSeriesColor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a Workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeSeriesColor.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the first chart
            Chart chart = sheet.Charts[0];

            //Get the second series
            ChartSerie cs = chart.Series[1];

            //Set the fill type
            cs.Format.Fill.FillType = ShapeFillType.SolidColor;

            //Change the fill color
            cs.Format.Fill.ForeColor = Color.Orange;

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
