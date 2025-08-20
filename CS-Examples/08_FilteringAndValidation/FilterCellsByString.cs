using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.AutoFilter;
using System;
using System.Windows.Forms;

namespace FilterCellsByString
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new Workbook object
            Workbook workbook = new Workbook();

            // Load an existing Excel file from the specified path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_6.xlsx");

            // Get the first sheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the range for filtering cells data, in this case, column D from row 1 to 19
            sheet.AutoFilters.Range = sheet.Range["D1:D19"];

            // Get the filter column for custom filtering
            FilterColumn filtercolumn = (FilterColumn)sheet.AutoFilters[0];

            // Apply a custom filter to display only cells starting with "South"
            sheet.AutoFilters.CustomFilter(filtercolumn, FilterOperatorType.Equal, "South*");

            // Apply the filters
            sheet.AutoFilters.Filter();

            // Save the filtered data to a new Excel file
            workbook.SaveToFile("filterCellsByString_result.xlsx", ExcelVersion.Version2013); 

            // Launch the file
            FileViewer("filterCellsByString_result.xlsx");

            // Dispose of the workbook object to release resources
            workbook.Dispose();
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
