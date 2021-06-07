using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace CreateTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateTable.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // Add a new List Object to the worksheet
            sheet.ListObjects.Create("table", sheet.Range[1, 1, 19, 5]);
            // Add Default Style to the table    
            sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleLight9;
      
            //Save to file
            string result = "CreateTable_out.xlsx";
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
