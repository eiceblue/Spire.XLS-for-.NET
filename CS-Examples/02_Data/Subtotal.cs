using Spire.Xls;
using System;
using System.Windows.Forms;

namespace Subtotal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load the workbook from file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Subtotal.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            
            // Select the range of data to be used for subtotals (in this case, columns A and B, rows 1 to 18)
            CellRange range = sheet.Range["A1:B18"];
            // Apply subtotals to the selected data using the "Sum" function
            sheet.Subtotal(range, 0, new int[] {1}, SubtotalTypes.Sum, true, false, true);

            // Specify the file name for the resulting workbook after applying subtotals
            String result = "Subtotal_Out.xlsx";

            // Save the modified workbook to the specified file in Excel 2010 format
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
