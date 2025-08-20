using System;
using System.Windows.Forms;
using Spire.Xls;

namespace EachWorksheetToDifferentPDF
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\EachWorksheetToDifferentPDFSample.xlsx");

            //Save each sheet to PDF
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                string FileName = sheet.Name + ".pdf";
                //Save the sheet to PDF
                sheet.SaveToPdf(FileName);

                //Launch the result file
                ExcelDocViewer(FileName);
            }

            // Dispose of the workbook object to release resources
            workbook.Dispose();

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
