using Spire.Xls;
using System;
using System.Windows.Forms;

namespace RemoveAllDigitalSignatures
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load an existing Excel document from file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WithDigitalSignature.xlsx");

            // Remove all digital signatures from the workbook
            workbook.RemoveAllDigitalSignatures();

            // Specify the output filename for the workbook
            String result = "RemoveAllDigitalSignatures.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

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
    }
}
