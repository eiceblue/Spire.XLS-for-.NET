using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

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
            //Load the document from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\EachWorksheetToDifferentPDFSample.xlsx");

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                string FileName = sheet.Name + ".pdf";
                //Save the sheet to PDF
                sheet.SaveToPdf(FileName);

                //Launch the result file
                ExcelDocViewer(FileName);
            }

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
