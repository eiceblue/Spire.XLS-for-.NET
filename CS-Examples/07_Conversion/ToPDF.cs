using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace ToPDF
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF.xlsx");

            // Set the ConverterSetting property to enable fitting sheets to page during PDF conversion
            workbook.ConverterSetting.SheetFitToPage = true;

            // Save the workbook as a PDF file with the name "sample.pdf"
            workbook.SaveToFile("sample.pdf", FileFormat.PDF);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("sample.pdf");
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
