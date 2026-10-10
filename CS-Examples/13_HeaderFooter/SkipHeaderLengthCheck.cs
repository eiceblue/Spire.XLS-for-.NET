using Spire.Xls;
using System;
using System.Windows.Forms;

namespace SkipHeaderLengthCheck
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Initialize a new Workbook instance
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Disable the validation for header/footer length to bypass limit checks
            workbook.ConverterSetting.ValidateHeaderFooterLength = false;

            // Define a very long string to intentionally exceed the default header length limit
            string fullLine = "══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════";

            // Set the long string as the left header content
            worksheet.PageSetup.LeftHeader = $"{fullLine}";

            // Specify the output file name for the PDF
            String result = "SkipHeaderLengthCheck.pdf";

            // Export the worksheet to a PDF file
            worksheet.SaveToPdf(result);

            // Dispose of the workbook to release resources
            workbook.Dispose();

            // View the document using a file viewer
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
