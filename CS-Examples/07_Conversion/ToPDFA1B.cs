using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace ToPDFA1B
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

            //Load an excel file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF_A1BExample.xlsx");

            //Convert excel to PDFA/1-B
            workbook.ConverterSetting.PdfConformanceLevel = Spire.Pdf.PdfConformanceLevel.Pdf_A1B;

            //Save the document and launch it
            workbook.SaveToFile("ToPDFA1B_result.pdf", FileFormat.PDF);

            FileViewer("ToPDFA1B_result.pdf");
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
