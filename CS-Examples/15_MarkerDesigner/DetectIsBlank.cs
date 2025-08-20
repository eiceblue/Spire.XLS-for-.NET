using System;
using System.Data;
using System.Windows.Forms;
using Spire.Xls;

namespace DetectIsBlank
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            // Create a workbook
			Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MarkerDesigner2.xlsx");

            // Create a DataSet
            DataSet ds = new DataSet();

            // Fill the DataSet from an XML file
            ds.ReadXml(@"..\..\..\..\..\..\Data\Data.xml");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Fill a DataTable using the "data" parameter in Marker Designer
            workbook.MarkerDesigner.AddDataTable("data", ds.Tables["data"]);
            workbook.MarkerDesigner.Apply();

            // Calculate all formulas in the workbook
            workbook.CalculateAllValue();

            // Save the modified workbook to a file named "DetectIsBlank.xlsx" using Excel 2013 format
            string output = "DetectIsBlank.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
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
