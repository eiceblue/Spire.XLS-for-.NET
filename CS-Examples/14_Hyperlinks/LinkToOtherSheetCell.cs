using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace LinkToOtherSheetCell
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

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];
          
            CellRange range = sheet.Range["A1"];

            //Add hyperlink in the range
            HyperLink hyperlink = sheet.HyperLinks.Add(range);

            //Set the link type
            hyperlink.Type = HyperLinkType.Workbook;

            //Set the display text
            hyperlink.TextToDisplay = "Link to Sheet2 cell C5";

            //Set the address
            hyperlink.Address = "Sheet2!C5";

            //Save and launch result file
            string result = "result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);
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
