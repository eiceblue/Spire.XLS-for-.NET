﻿using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace ConsolidationFunctions
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PivotTableExample.xlsx");
            //Get the sheet in which the pivot table is located
            Worksheet sheet = workbook.Worksheets["PivotTable"];

            XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

            //Apply Average consolidation function to first data field
            pt.DataFields[0].Subtotal = SubtotalTypes.Average;
  
            //Apply Max consolidation function to second data field
            pt.DataFields[1].Subtotal = SubtotalTypes.Max;         
          
            pt.CalculateData();

            String result = "ConsolidationFunctions_result.xlsx";

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
