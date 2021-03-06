﻿using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FindCellsWithStyleName
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load the document from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the cell style name
            string styleName = sheet.Range["A1"].CellStyleName;

            CellRange ranges = sheet.AllocatedRange;
            foreach (CellRange cc in ranges)
            {
                //Find the cells which have the same style name
                if (cc.CellStyleName == styleName)
                {
                    //Set value
                    cc.Value = "Same style";
                }
            }

            string result = "FindCellsWithStyleName_result.xlsx";
            //Save and launch result file
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
