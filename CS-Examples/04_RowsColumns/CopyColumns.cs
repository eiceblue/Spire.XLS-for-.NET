using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace CopyColumns
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

            //Load an excel file including pivot table
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Copying.xls");

            Worksheet sheet1 = workbook.Worksheets[0];
            Worksheet sheet2 = workbook.Worksheets[1];

            //Copy the first column to the third column in the same sheet
            sheet1.Copy(sheet1.Columns[0],sheet1.Columns[2],true,true,true);

            //Copy the first column to the second column in the different sheet
            sheet1.Copy(sheet1.Columns[0], sheet2.Columns[1], true, true, true);

            String result = "CopyColumns_result.xlsx";

            //Save to file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

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
