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
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\TrackChanges.xlsx");

            //Accept the changes or reject the changes.
            //workbook.AcceptAllTrackedChanges();
            workbook.RejectAllTrackedChanges();

            //Save to file.
            String outputFile = "AcceptOrRejectTrackedChanges.xlsx";       
            workbook.SaveToFile(outputFile, FileFormat.Version2013);

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
