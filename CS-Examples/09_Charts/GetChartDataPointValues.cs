using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

using System.Text;
using System.IO;

namespace GetChartDataPointValues
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            //Load the document from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartToImage.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the chart
            Chart chart = sheet.Charts[0];

            //Get the first series of the chart
            ChartSerie cs = chart.Series[0];

            foreach (CellRange cr in cs.Values)
            {
                sb.Append(cr.RangeAddress + "\r\n");

                //Get the data point value
                sb.Append("The value of the data point is " + cr.Value + "\r\n");
            }

            string result = "result.txt";
            //Save and launch result file
            File.WriteAllText(result, sb.ToString());
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
