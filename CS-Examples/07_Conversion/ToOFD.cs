using System;
using System.Windows.Forms;

using Spire.Xls;

namespace ToOFD
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
            
            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToOFD.xlsx");
            //Save to ofd file
            workbook.SaveToFile("result.ofd", FileFormat.OFD);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("result.ofd");
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
