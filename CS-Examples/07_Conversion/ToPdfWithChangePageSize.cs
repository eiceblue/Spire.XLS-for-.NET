using System;
using System.Windows.Forms;
using Spire.Xls;


namespace ToPdfWithChangePageSize
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                //Change the page size
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            }

            //Save the result file
            string result = "result.pdf";
            workbook.SaveToFile(result, FileFormat.PDF);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
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
