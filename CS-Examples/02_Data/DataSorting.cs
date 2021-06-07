using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace DataSorting
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataSorting.xls");

            Worksheet worksheet = workbook.Worksheets[0];


            workbook.DataSorter.SortColumns.Add(2, OrderBy.Ascending);
            workbook.DataSorter.SortColumns.Add(3, OrderBy.Ascending);
            
            workbook.DataSorter.Sort(worksheet["A1:E19"]);

            string result = "DataSorting_out.xlsx";
			workbook.SaveToFile(result,ExcelVersion.Version2013);

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

        private void Form1_Load(object sender, EventArgs e)
        {

        }

	}
}
