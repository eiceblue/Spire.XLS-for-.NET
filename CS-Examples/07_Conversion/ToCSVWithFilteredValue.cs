using Spire.Xls;
using System;
using System.Windows.Forms;

namespace ToCSVWithFilteredValue
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

            // Load  excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AutofilterSample.xlsx");
       
            // Convert to CSV file with filtered value
            workbook.Worksheets[0].SaveToFile("ToCSVWithFilteredValue.csv", ";", false);

            // Convert to CSV stream
            //worksheet.SaveToStream(Stream stream, string separator, bool retainHiddenData);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the document
            FileViewer("ToCSVWithFilteredValue.csv");
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
