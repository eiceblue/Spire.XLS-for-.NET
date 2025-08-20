using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SaveFiles
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

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

            // Save in Excel 97-2003 format
            workbook.SaveToFile("result.xls",ExcelVersion.Version97to2003);

            // Save in Excel2010 xlsx format
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010);

            // Save in XLSB format
            workbook.SaveToFile("result.xlsb", ExcelVersion.Xlsb2010);

            // Save in ODS format
            workbook.SaveToFile("result.ods", ExcelVersion.ODS);

            // Save in PDF format
            workbook.SaveToFile("result.pdf", FileFormat.PDF);

            // Save in XML format
            workbook.SaveToFile("result.xml",FileFormat.XML);

            // Save in XPS format
            workbook.SaveToFile("result.xps", FileFormat.XPS);

            // Dispose of the workbook object to release resources 
            workbook.Dispose();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
