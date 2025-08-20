using System;
using System.Windows.Forms;
using Spire.Xls;

namespace FitWidthWhenConvertToPDF
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a Workbook
            Workbook workbook = new Workbook();

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Auto fit page height
                sheet.PageSetup.FitToPagesTall = 0;
                // Fit one page width
                sheet.PageSetup.FitToPagesWide = 1;
            }

            // Save result file
            string result = "result.pdf";
            workbook.SaveToFile(result, FileFormat.PDF);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
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
