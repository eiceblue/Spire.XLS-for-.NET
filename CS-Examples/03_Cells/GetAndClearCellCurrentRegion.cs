using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core;

namespace GetAndClearCellCurrentRegion
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new Workbook object.
            Workbook workbook = new Workbook();

            // Load an existing Excel file from the specified path.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_10.xlsx");

            // Get the first worksheet from the workbook.
            Worksheet sheet = workbook.Worksheets[0];

            // Get the current region of the cell starting from cell A1 and clear its contents.
            IXLSRange xlRange = sheet.Range["A1"].CurrentRegion;
            foreach (CellRange range in xlRange)
            {
                range.ClearAll();
            }

            // Specify the filename for the resulting Excel file.
            string result = "CellCurrentRegion_result.xlsx";

            // Save the modified workbook to a new file with the specified filename and Excel version.
            workbook.SaveToFile(result, ExcelVersion.Version2016);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the MS Excel file.
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
