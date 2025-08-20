using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace CustomPivotTableFieldName
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

            // Load an excel file including pivot table   
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CustomPivotTableFieldName.xlsx");        

            // Get the sheet in which the pivot table is located
            Worksheet sheet = workbook.Worksheets["PivotTable"];

            // Access the first pivot table in the worksheet
            XlsPivotTable pivotTable = sheet.PivotTables[0] as XlsPivotTable;

            // Set a custom name for the row field
            pivotTable.RowFields[0].CustomName = "custom_rowName";

            // Set a custom name for the column field
            pivotTable.ColumnFields[0].CustomName = "custom_colName";

            // Set a custom name for the data field
            pivotTable.DataFields[0].CustomName = "custom_DataName";

            // Calculate the pivot table data
            pivotTable.CalculateData();

            // Specify the filename for the resulting workbook
            string result = "CustomPivotTableFieldName_result.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
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

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
