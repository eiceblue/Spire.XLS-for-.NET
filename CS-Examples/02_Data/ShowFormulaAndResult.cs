using System;
using System.Data;
using System.Windows.Forms;
using Spire.Xls;

namespace ShowFormulaAndResult
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		//Formula
		private void btnRun_Click(object sender, EventArgs e)
		{
            //Create a workbook
			Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FormulasSample.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

			//Show formula
            DataTable dt = sheet.ExportDataTable(sheet.AllocatedRange, false, false);

			//Show in DataGridView
            this.dataGridView1.DataSource = dt;

            // Dispose of the workbook object to release resources
            workbook.Dispose();
        }

		//Result
        private void btnClose_Click(object sender, EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FormulasSample.xlsx");

            Worksheet sheet = workbook.Worksheets[0];
			//Show result
            DataTable dt = sheet.ExportDataTable(sheet.AllocatedRange, false, true);
			////Show in DataGridView
            this.dataGridView1.DataSource = dt;

            // Dispose of the workbook object to release resources
            workbook.Dispose();
        }

	}
}
