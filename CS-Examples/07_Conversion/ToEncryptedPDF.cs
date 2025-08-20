using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls.Pdf.Security;
using Spire.Xls;

namespace ToEncryptedPdf
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook
            Workbook workbook = new Workbook();
            
            // Load a excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF.xlsx");

            // Set open and permission password to encrypt converted pdf
            workbook.ConverterSetting.PdfSecurity.Encrypt("123","456", PdfPermissionsFlags.Print, PdfEncryptionKeySize.Key128Bit);

            // Convert excel to pdf
            workbook.SaveToFile("sample.pdf", FileFormat.PDF);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
            ExcelDocViewer("sample.pdf");
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
