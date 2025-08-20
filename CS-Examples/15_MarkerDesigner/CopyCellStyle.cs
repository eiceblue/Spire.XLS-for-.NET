using System;
using System.Data;
using System.Windows.Forms;
using Spire.Xls;

namespace CopyCellStyle
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MarkerDesigner1.xlsx");

            // Create a DataTable
            DataTable dt = new DataTable("data");

            // Define columns in the DataTable
            dt.Columns.Add(new DataColumn("name", typeof(string)));
            dt.Columns.Add(new DataColumn("age", typeof(int)));

            // Add three rows to the DataTable
            DataRow drName1 = dt.NewRow();
            DataRow drName2 = dt.NewRow();
            DataRow drName3 = dt.NewRow();

            drName1["name"] = "John";
            drName1["age"] = 15;
            drName2["name"] = "Jess";
            drName2["age"] = 22;
            drName3["name"] = "Alan";
            drName3["age"] = 36;

            dt.Rows.Add(drName1);
            dt.Rows.Add(drName2);
            dt.Rows.Add(drName3);

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Fill the DataTable using the "data" parameter in Marker Designer
            workbook.MarkerDesigner.AddDataTable("data", dt);
            workbook.MarkerDesigner.Apply();

            // Specify the output file name for the modified workbook
            string output = "CopyCellStyle.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the file
            ExcelDocViewer(output);
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
