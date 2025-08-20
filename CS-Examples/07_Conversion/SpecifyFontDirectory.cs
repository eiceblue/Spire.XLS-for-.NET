using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SpecifyFontDirectory
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a workbook
            Workbook workbook = new Workbook();
            
            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDFSample.xlsx");

            // Specify font directory
            workbook.CustomFontFileDirectory = new string[] { (@"..\..\..\..\..\..\Data\Font") };

            // Save to pdf file
            workbook.SaveToFile("result.pdf", FileFormat.PDF);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("result.pdf");
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
