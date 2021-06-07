using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CopyVisibleSheets
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

            //Load a csv file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CopyVisibleSheets.xlsx");

            //Create a new workbook
            Workbook workbookNew = new Workbook();
            workbookNew.Version = ExcelVersion.Version2013;
            workbookNew.Worksheets.Clear();

            //Loop through the worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                //Judge if the worksheet is visible
                if (sheet.Visibility == WorksheetVisibility.Visible)
                {
                    //Copy the sheet to new workbook
                    string name = sheet.Name;
                    workbookNew.Worksheets.AddCopy(sheet);
                }
            }

            //Save the Excel file
            string result = "CopyVisibleSheets_out.xlsx";
            workbookNew.SaveToFile(result, ExcelVersion.Version2013);
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
        private void btnClose_Click_1(object sender, EventArgs e)
        {
            Close();
        }
    }
}
