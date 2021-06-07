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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DigitalSignature.xlsx");
            //Add certificate
            String inputFile_pfx = @"..\..\..\..\..\..\Data\gary.pfx";            
            X509Certificate2 cert = new X509Certificate2(inputFile_pfx, "e-iceblue");
            //Add signature
            DateTime certtime = new DateTime(2020, 7, 1, 7, 10, 36);
            IDigitalSignatures dsc = workbook.AddDigitalSignature(cert, "e-iceblue", certtime);

            String result = "AddDigitalSignature.xlsx";

            //Save to file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

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
    }
}
