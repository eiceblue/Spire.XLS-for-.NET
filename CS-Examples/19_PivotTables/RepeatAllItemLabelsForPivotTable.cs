using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace RepeatAllItemLabelsForPivotTable
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new instance of the Workbook class
            Workbook workbook = new Workbook();

            // Load the workbook from the specified file path 
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\RepeatAllItemLabelsForPivotTable.xlsx");

            // Iterate through each pivot table in the "Pivot" worksheet
            foreach (XlsPivotTable pt in workbook.Worksheets["Pivot"].PivotTables)
            {
                // Set the RepeatAllItemLabels property to true for the pivot table
                pt.Options.RepeatAllItemLabels = true;

                // Calculate the data for the pivot table
                pt.CalculateData();

                // Refresh the cache for the pivot table
                pt.Cache.IsRefreshOnLoad = true;
            }

            // Define the output file name for the modified workbook
            String result = "RepeatAllItemLabelsForPivotTable_output.xlsx";

            // Save the modified workbook to the specified file path with the specified Excel version
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose the workbook instance to release resources
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

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
