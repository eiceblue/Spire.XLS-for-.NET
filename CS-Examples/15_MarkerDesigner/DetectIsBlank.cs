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
            //Create a workbook
			Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MarkerDesigner2.xlsx");

            DataSet ds = new DataSet();
            //Fill dataset from XML file
            ds.ReadXml(@"..\..\..\..\..\..\Data\Data.xml");

            Worksheet sheet = workbook.Worksheets[0];
            //Fill DataTable
            workbook.MarkerDesigner.AddDataTable("data", ds.Tables["data"]);
            workbook.MarkerDesigner.Apply();

            //Calculate formulas
            workbook.CalculateAllValue();

            //Save the document
            string output = "DetectIsBlank.xlsx";
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
