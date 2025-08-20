using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace MarkerDesigner
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MarkerDesigner.xls");

            // Get the DataTable from the DataSource of dataGrid1
            DataTable dt = (DataTable)dataGrid1.DataSource;

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Fill a parameter named "Variable1" with the value 1234.5678
            workbook.MarkerDesigner.AddParameter("Variable1", 1234.5678);

            // Fill a DataTable named "Country" with the data from dt
            workbook.MarkerDesigner.AddDataTable("Country", dt);
            workbook.MarkerDesigner.Apply();

            // AutoFit rows and columns to adjust their sizes based on content
            sheet.AllocatedRange.AutoFitRows();
            sheet.AllocatedRange.AutoFitColumns();

            // Specify the output file name for the modified workbook
            String result = "Output_MarkerDesigner.xlsx";

            // Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
		}

		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();

            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MarkerDesigner-DataSample.xls");
			//Initailize worksheet
			Worksheet sheet = workbook.Worksheets[0];

            this.dataGrid1.DataSource = sheet.ExportDataTable();
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }


	}
}
