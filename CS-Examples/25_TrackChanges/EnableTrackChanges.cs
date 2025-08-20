using Spire.Xls;
using System;
using System.Windows.Forms;

namespace EnableTrackChanges
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load an existing Excel file from the specified path
             workbook.LoadFromFile(@"..\..\..\..\..\..\Data\textAlign.xlsx");

            //Enable track changes 
            workbook.TrackedChanges = true;

            // Specify the filename for the resulting Excel file
            String result = "output.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object
            workbook.Dispose();

            // View the document using a file viewer
            FileViewer(result);

            this.Close();
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
