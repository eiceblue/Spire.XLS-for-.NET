using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core.Spreadsheet.Collections;

namespace ColorsAndPalette
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Adding Orchid color to the palette at 60th index
            workbook.ChangePaletteColor(Color.Orchid, 60);

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            CellRange cell = sheet.Range["B2"];
            cell.Text = "Welcome to use Spire.XLS";

            //Set the Orchid (custom) color to the font
            cell.Style.Font.Color = Color.Orchid;
            cell.Style.Font.Size = 20;
            cell.AutoFitColumns();
            cell.AutoFitRows();

            //Save to file
            String result = "ColorsAndPalette_result.xlsx";
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
