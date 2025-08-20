using System;
using System.Windows.Forms;

using Spire.Xls;

namespace LoadSaveEtAndETT
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //create a workbook
            Workbook workbook = new Workbook();

            //load .et or .ett file 
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Sample-et.et");
            //workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Sample-ett.ett");

            //save to .et or .ett file
            workbook.SaveToFile("result.et", FileFormat.ET);
            //workbook.SaveToFile("result.ett", FileFormat.ETT);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //view the document
            ExcelDocViewer("result.et");
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
