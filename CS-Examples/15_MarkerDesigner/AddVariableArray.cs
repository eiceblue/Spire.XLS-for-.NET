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
            // Create a workbook
			Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set marker designer field in cell A1
            sheet.Range["A1"].Value = "&=Array";

            // Fill an array using the "Array" parameter
            workbook.MarkerDesigner.AddArray("Array", new string[] { "Spire.Xls", "Spire.Doc", "Spire.PDF", "Spire.Presentation", "Spire.Email" });
            workbook.MarkerDesigner.Apply();
            workbook.CalculateAllValue();

            // AutoFit rows and columns to adjust their sizes based on content
            sheet.AllocatedRange.AutoFitRows();
            sheet.AllocatedRange.AutoFitColumns();

            // Specify the output file name for the modified workbook
            string output = "AddVariableArray.xlsx";
            
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
