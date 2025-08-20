using System;
using System.Windows.Forms;
using Spire.Xls;

namespace ToXPS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load a file from the specified path into the workbook
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToXPS.xlsx");

            // Save the workbook as an XPS file with the name "ToXPS.xps" using the Spire.Xls library's XPS file format
            workbook.SaveToFile("ToXPS.xps", Spire.Xls.FileFormat.XPS);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("ToXPS.xps");
        }

        private void ExcelDocViewer(string fileName)
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
