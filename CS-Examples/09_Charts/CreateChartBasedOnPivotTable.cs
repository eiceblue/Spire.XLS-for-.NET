using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateChartBasedOnPivotTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load an excel file including pivot table
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PivotTable.xlsx");
          
            //Get the sheet in which the pivot table is located
            Worksheet sheet = workbook.Worksheets[0];

            XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

            workbook.Worksheets[1].Charts.Add(ExcelChartType.BarClustered, pt);

            //Save the document
            string output = "CreateChartBasedOnPivotTable.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            //View the document
            FileViewer(output);
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
