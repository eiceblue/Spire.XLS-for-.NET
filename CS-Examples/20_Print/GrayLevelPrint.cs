using System;
using System.Windows.Forms;
using Spire.Xls;

namespace GrayLevelPrint
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_3.xlsx");

            // Set the GrayLevelForPrint to true
            workbook.ConverterSetting.GrayLevelForPrint = true;

            // Print this document
            workbook.PrintDocument.Print();

            // Dispose of the workbook object to release resources
            workbook.Dispose();
          
        }
        private void OutputViewer(string fileName)
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
