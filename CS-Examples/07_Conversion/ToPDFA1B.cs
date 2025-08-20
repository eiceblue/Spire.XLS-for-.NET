using Spire.Xls;
using Spire.Xls.Pdf;
using System;
using System.Windows.Forms;


namespace ToPDFA1B
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

            // Load an excel file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF_A1BExample.xlsx");

            // Convert excel to PDFA/1-B
            workbook.ConverterSetting.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;

            // Save the document and launch it
            workbook.SaveToFile("ToPDFA1B_result.pdf", FileFormat.PDF);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            FileViewer("ToPDFA1B_result.pdf");
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
