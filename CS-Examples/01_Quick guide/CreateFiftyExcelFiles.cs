using Spire.Xls;
using System;
using System.Windows.Forms;

namespace CreateFiftyExcelFiles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Get the current date and time
            DateTime start = DateTime.Now;

            // Create 50 workbooks with 5 sheets each
            for (int n = 0; n < 50; n++)
            {
                // Create a new workbook object
                Workbook workbook = new Workbook();

                // Create 5 empty sheets in the workbook
                workbook.CreateEmptySheets(5);

                // Iterate through each sheet in the workbook
                for (int i = 0; i < 5; i++)
                {
                    // Get the reference to the current sheet
                    Worksheet sheet = workbook.Worksheets[i];

                    // Set the name of the sheet based on the index
                    sheet.Name = "Sheet" + i.ToString();

                    // Populate the sheet with data
                    for (int row = 1; row <= 150; row++)
                    {
                        for (int col = 1; col <= 50; col++)
                        {
                            // Set the cell value to a combination of the current row and column numbers
                            sheet.Range[row, col].Text = "row" + row.ToString() + " col" + col.ToString();
                        }
                    }
                }

                // Specify the filename for the resulting Excel file, using the iteration number
                workbook.SaveToFile("Workbook" + n + ".xlsx", ExcelVersion.Version2010);

                // Dispose of the workbook object
                workbook.Dispose();
            }

            // Get the current date and time after the operation
            DateTime end = DateTime.Now;

            // Calculate the time taken by subtracting the start time from the end time
            TimeSpan time = end - start;

            // Show a message box with the success message and the time consumed
            MessageBox.Show("50 File(s) have been created successfully! \n" + "Time consumed (Seconds): " + time.TotalSeconds.ToString());
          
        }


        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
