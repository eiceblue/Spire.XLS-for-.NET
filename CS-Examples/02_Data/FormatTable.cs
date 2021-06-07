using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace FormatTable
{
    public partial class Form1 :Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FormatTable.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            //Add Default Style to the table
            sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9;
            //Show Total
            sheet.ListObjects[0].DisplayTotalRow = true;
            //Set calculation type
            sheet.ListObjects[0].Columns[0].TotalsRowLabel = "Total";
            sheet.ListObjects[0].Columns[1].TotalsCalculation = ExcelTotalsCalculation.None;
            sheet.ListObjects[0].Columns[2].TotalsCalculation = ExcelTotalsCalculation.None;
            sheet.ListObjects[0].Columns[3].TotalsCalculation = ExcelTotalsCalculation.Sum;
            sheet.ListObjects[0].Columns[4].TotalsCalculation = ExcelTotalsCalculation.Sum;
          
            sheet.ListObjects[0].ShowTableStyleRowStripes = true;
            
            sheet.ListObjects[0].ShowTableStyleColumnStripes = true;
            workbook.SaveToFile("Sample.xlsx", ExcelVersion.Version2010);
            ExcelDocViewer("Sample.xlsx");
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
