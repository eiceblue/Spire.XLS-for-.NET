using Spire.Xls;
using System;
using System.Windows.Forms;

namespace CreateAnExcelWithOneSheet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Record the current time as the starting point
            DateTime start = DateTime.Now;

            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Create an empty sheet in the workbook
            workbook.CreateEmptySheets(1);

            // Get the reference to the first sheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the sheet with data
            for (int row = 1; row <= 10000; row++)
            {
                for (int col = 1; col <= 30; col++)
                {
                    // Set the cell value to a combination of the current row and column numbers
                    sheet.Range[row, col].Text = row.ToString() + "," + col.ToString();
                }
            }

            // Specify the filename for the resulting Excel file
            String result = "CreateAnExcelWithOneSheet_result.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object
            workbook.Dispose();

            // Record the current time as the ending point
            DateTime end = DateTime.Now;

            // Calculate the time taken to create the file and display it as a message box
            TimeSpan time = end - start;
            MessageBox.Show("File has been created successfully! \n" + "Time consumed (Seconds): " + time.TotalSeconds.ToString());

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
