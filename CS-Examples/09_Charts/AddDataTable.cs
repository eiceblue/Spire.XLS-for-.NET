using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace AddDataTable
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddDataTable.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the first chart
            Chart chart = sheet.Charts[0];
           
            chart.HasDataTable = true;

            //Save and Launch
            workbook.SaveToFile("Output.xlsx", FileFormat.Version2010);
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
