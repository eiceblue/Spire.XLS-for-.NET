using System;
using System.Windows.Forms;
using Spire.Xls;

namespace ToUOS
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToUOS.xlsx");

            //Save to uos file
            workbook.SaveToFile("result.uos", FileFormat.UOS);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file 
            ExcelDocViewer("result.uos");
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
