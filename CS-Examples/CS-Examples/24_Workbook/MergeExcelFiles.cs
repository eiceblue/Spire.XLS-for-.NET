using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MergeExcelFiles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> files = new List<string>();
            files.Add(@"..\..\..\..\..\..\Data\MergeExcelFiles-1.xlsx");
            files.Add(@"..\..\..\..\..\..\Data\MergeExcelFiles-2.xls");
            files.Add(@"..\..\..\..\..\..\Data\MergeExcelFiles-3.xlsx");

            Workbook newbook = new Workbook();
            newbook.Version = ExcelVersion.Version2013;
            //Clear all worksheets
            newbook.Worksheets.Clear();

            //Create a workbook
            Workbook tempbook = new Workbook();

            foreach (string file in files)
            {
                //Load the file
                tempbook.LoadFromFile(file);
                foreach (Worksheet sheet in tempbook.Worksheets)
                {
                    //Copy every sheet in a workbook
                    newbook.Worksheets.AddCopy(sheet,WorksheetCopyType.CopyAll);
                }
            }

            //Save the file
            newbook.SaveToFile("MergeExcelFiles.xlsx", ExcelVersion.Version2010);
            ExcelDocViewer("MergeExcelFiles.xlsx");
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
