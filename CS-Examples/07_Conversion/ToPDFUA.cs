using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace ToPDFUA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

            // Convert excel to PDF/UA
            workbook.ConverterSetting.PdfConformanceLevel = Spire.Xls.Pdf.PdfConformanceLevel.Pdf_UA1;

            // Specify the output file name for the EXCEL result
            string result = "ToPDFUA_out.pdf";

            // Save the workbook as a PDF file with the name "sample.pdf"
            workbook.SaveToFile(result, FileFormat.PDF);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
