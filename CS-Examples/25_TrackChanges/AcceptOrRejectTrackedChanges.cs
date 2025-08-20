using Spire.Xls;
using System;
using System.Windows.Forms;

namespace AcceptOrRejectTrackedChanges
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

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\TrackChanges.xlsx");

            // Accept the changes or reject the changes.
            //workbook.AcceptAllTrackedChanges();
            workbook.RejectAllTrackedChanges();

            // Save to file.
            string outputFile = "AcceptOrRejectTrackedChanges.xlsx";       
            workbook.SaveToFile(outputFile, FileFormat.Version2013);

            // Dispose of the workbook object to release resources 
            workbook.Dispose();

            //View the document
            FileViewer(outputFile);
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
