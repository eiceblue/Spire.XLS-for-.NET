using Spire.Xls;
using System;
using System.Windows.Forms;

namespace CreateAnExcelWithFiveSheets
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

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Create five empty sheets in the workbook
            workbook.CreateEmptySheets(5);

            // Iterate over each sheet in the workbook
            for (int i = 0; i < 5; i++)
            {
                // Get the current sheet
                Worksheet sheet = workbook.Worksheets[i];

                // Set the name of the sheet using the index
                sheet.Name = "Sheet" + i.ToString();

                // Populate the sheet with data in a grid-like pattern
                for (int row = 1; row <= 150; row++)
                {
                    for (int col = 1; col <= 50; col++)
                    {
                        // Set the text in each cell of the sheet using the row and column numbers
                        sheet.Range[row, col].Text = "row" + row.ToString() + " col" + col.ToString();
                    }
                }
            }

            // Specify the file name for saving the workbook
            String result = "CreateAnExcelWithFiveSheets_result.xlsx";

            // Save the workbook to a file with Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object
            workbook.Dispose();

            // Record the current time as the ending point
            DateTime end = DateTime.Now;

            // Calculate the time taken to create the file and display it as a message box
            TimeSpan time = end - start;
            MessageBox.Show("File has been created successfully! \n" + "Time consumed (Seconds): " + time.TotalSeconds.ToString());
            
            // View the document
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
