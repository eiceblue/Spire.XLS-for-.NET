using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace FormatDataField
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FormatDataField.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            // Access the PivotTable
            XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;
            // Access the data field.
            PivotDataField pivotDataField = pt.DataFields[0];
            // Set data display format
            pivotDataField.ShowDataAs = PivotFieldFormatType.PercentageOfColumn;

            String result = "FormatDataField_output.xlsx";
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
