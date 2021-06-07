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

            // Create Students DataTable
            DataTable dt = new DataTable("data");

            // Define a field in it
            dt.Columns.Add(new DataColumn("name", typeof(string)));
            dt.Columns.Add(new DataColumn("age", typeof(int)));

            // Add three rows to it
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

            Worksheet sheet = workbook.Worksheets[0];
            //Fill DataTable
            workbook.MarkerDesigner.AddDataTable("data", dt);
            workbook.MarkerDesigner.Apply();

            //Save the document
            string output = "CopyCellStyle.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

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
