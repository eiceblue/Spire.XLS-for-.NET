using Spire.Xls;
using System;
using System.Windows.Forms;

namespace ToPdfSimply
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

            // Load a excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF.xlsx");

            // Convert excel to pdf
            string result = "ToPdfSimply_result.pdf";
            workbook.SaveToFile(result, FileFormat.PDF);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the document
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
