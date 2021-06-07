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

namespace GetCategoryLabels
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

            //Create a workbook
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampeB_4.xlsx");

            Worksheet sheet = workbook.Worksheets[0];

            //Get the chart
            Chart chart = sheet.Charts[0];

            //Get the cell range of the category labels
            CellRange cr = chart.PrimaryCategoryAxis.CategoryLabels;
            foreach (var cell in cr)
            {
                sb.Append(cell.Value + "\r\n");
            }

            //Save and launch result file
            string result = "result.txt";
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
