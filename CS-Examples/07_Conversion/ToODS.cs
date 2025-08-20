using Spire.Xls;
using System;
using System.Windows.Forms;

namespace ToODS
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToODS.xlsx");
   
            // Convert to ODS file
            workbook.SaveToFile("Result.ods", FileFormat.ODS);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // view the document
            ExcelDocViewer("Result.ods");
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
