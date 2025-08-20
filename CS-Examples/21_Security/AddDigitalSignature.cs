using Spire.Xls;
using Spire.Xls.Core.MergeSpreadsheet.Interfaces;
using System;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;

namespace AddDigitalSignature
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

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DigitalSignature.xlsx");

            // Add a digital certificate for signing the workbook
            String inputFile_pfx = @"..\..\..\..\..\..\Data\gary.pfx";            
            X509Certificate2 cert = new X509Certificate2(inputFile_pfx, "e-iceblue");

            // Specify the date and time for the digital signature
            DateTime certtime = new DateTime(2020, 7, 1, 7, 10, 36);

            // Add a digital signature to the workbook using the provided certificate, signer name ("e-iceblue"), and signature timestamp
            IDigitalSignatures dsc = workbook.AddDigitalSignature(cert, "e-iceblue", certtime);

            // Specify the output filename for the signed workbook
            String result = "AddDigitalSignature.xlsx";

            // Save the workbook with the added digital signature to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

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
