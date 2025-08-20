using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MakeCellActive
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

            // Load file froom disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\templateAz.xlsx");

            // Get the 2nd sheet
            Worksheet sheet =  workbook.Worksheets[1];

            // Set the 2nd sheet as an active sheet.
            sheet.Activate();

            // Set B2 cell as an active cell in the worksheet.
            sheet.SetActiveCell(sheet.Range["B2"]);

            // Set the B column as the first visible column in the worksheet.
            sheet.FirstVisibleColumn = 1;

            // Set the 2nd row as the first visible row in the worksheet.
            sheet.FirstVisibleRow = 1;

            // Specify the name for the resulting Excel file
            String result = "MakeCellActive_result.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the document
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
