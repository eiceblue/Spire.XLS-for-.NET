using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace PivotTableLayout
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load the file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PivotTable.xlsx");

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first PivotTable
            XlsPivotTable xlsPivotTable = (XlsPivotTable)worksheet.PivotTables[0];

            // Set the PivotTable layout type
            xlsPivotTable.Options.ReportLayout = PivotTableLayoutType.Tabular;

            // Save to file
            String result = "PivotLayoutTabular_output.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // View the document
            FileViewer(result);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
