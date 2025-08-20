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
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Load an Excel file into the workbook
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataSorting.xls");

            // Get the first worksheet from the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a sorting column: column 2 in ascending order
            workbook.DataSorter.SortColumns.Add(2, OrderBy.Ascending);

            // Add another sorting column: column 3  in ascending order
            workbook.DataSorter.SortColumns.Add(3, OrderBy.Ascending);

            // Perform the sorting operation on the specified range: A1 to E19
            workbook.DataSorter.Sort(worksheet["A1:E19"]);

            // Set the output file name
            string result = "DataSorting_out.xlsx";

            // Save the sorted data to an Excel file (in this case, using the 2013 version)
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object
            workbook.Dispose();

            // View the output file
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
