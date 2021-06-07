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

namespace ChangeMajorGridlines
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

            //Get the chart
            Chart chart = sheet.Charts[0];

            //Change the color of major gridlines
            chart.PrimaryValueAxis.MajorGridLines.LineProperties.Color = Color.Red;

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
