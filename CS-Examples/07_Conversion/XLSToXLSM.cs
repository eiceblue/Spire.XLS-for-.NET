using System;
using System.Windows.Forms;
using Spire.Xls;

namespace XLSToXLSM
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MacroSample.xls",ExcelVersion.Version97to2003);

            //Save the workbook as a new XLSM file
            string output = "XLSToXLSM.xlsm";
			workbook.SaveToFile(output);

            //Launch the file
			ExcelDocViewer(output);
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
