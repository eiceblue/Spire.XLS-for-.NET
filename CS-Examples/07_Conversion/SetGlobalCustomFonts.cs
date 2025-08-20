using Spire.Xls;
using System;
using System.Windows.Forms;

namespace SetGlobalCustomFonts
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Set custom font directory
            string[] fontPath = { @"..\..\..\..\..\..\Data\fonts" };
      

            // Create a new workbook object
            Workbook workbook = new Workbook();
	    Workbook.SetGlobalCustomFontsFolders(fontPath);

            // Load an existing Excel file from the specified path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SpecialFont.xlsx");

            // Save the workbook to PDF 
            String result = "output.pdf";
            workbook.SaveToFile(result, FileFormat.PDF);

            // Dispose of the workbook object
            workbook.Dispose();

            // View the document using a file viewer
            FileViewer(result);

            this.Close();
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
