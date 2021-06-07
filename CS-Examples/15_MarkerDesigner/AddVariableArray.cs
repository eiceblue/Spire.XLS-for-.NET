using System;
using System.Windows.Forms;
using Spire.Xls;

namespace AddVariableArray
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

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set marker designer field in cell A1
            sheet.Range["A1"].Value = "&=Array";

            //Fill Array
            workbook.MarkerDesigner.AddArray("Array", new string[] { "Spire.Xls", "Spire.Doc", "Spire.PDF", "Spire.Presentation", "Spire.Email" });
            workbook.MarkerDesigner.Apply();
            workbook.CalculateAllValue();

            //AutoFit
            sheet.AllocatedRange.AutoFitRows();
            sheet.AllocatedRange.AutoFitColumns();

            //Save the document
            string output = "AddVariableArray.xlsx";
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
