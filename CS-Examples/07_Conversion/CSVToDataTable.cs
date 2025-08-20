using System;
using System.Windows.Forms;
using Spire.Xls;

namespace CSVToDataTable
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

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CSVSample.csv", ",");

            //Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //Export to datatable
            System.Data.DataTable dataTable = worksheet.ExportDataTable();

            //Show in data grid
            this.dataGridView1.DataSource = dataTable;

            // Dispose of the workbook object to release resources
            workbook.Dispose();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
