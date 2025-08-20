using Spire.Xls;
using System;
using System.Windows.Forms;

namespace ChangeDataSource
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

            //Load an excel file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeDataSource.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range of cells to be used as the new data source
            CellRange range = sheet.Range["A1:C15"];

            // Get the first pivot table from the second worksheet
            PivotTable table = workbook.Worksheets[1].PivotTables[0] as PivotTable;

            // Change the data source of the pivot table to the new range
            table.ChangeDataSource(range);

            // Disable automatic refresh of the pivot table cache on load
            table.Cache.IsRefreshOnLoad = false;

            // Specify the filename for the resulting workbook
            string result = "ChangeDataSource_result.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
            FileViewer(result);
         
        }

        private void FileViewer(string fileName)
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
