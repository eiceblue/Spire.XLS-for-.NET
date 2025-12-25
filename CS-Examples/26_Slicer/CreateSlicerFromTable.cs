using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateSlicerFromTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Workbook wb = new Workbook();
            Worksheet worksheet = wb.Worksheets[0];
            worksheet.Range["A1"].Value = "fruit";
            worksheet.Range["A2"].Value = "grape";
            worksheet.Range["A3"].Value = "blueberry";
            worksheet.Range["A4"].Value = "kiwi";
            worksheet.Range["A5"].Value = "cherry";
            worksheet.Range["A6"].Value = "grape";
            worksheet.Range["A7"].Value = "blueberry";
            worksheet.Range["A8"].Value = "kiwi";
            worksheet.Range["A9"].Value = "cherry";

            worksheet.Range["B1"].Value = "year";
            worksheet.Range["B2"].Value2 = 2020;
            worksheet.Range["B3"].Value2 = 2020;
            worksheet.Range["B4"].Value2 = 2020;
            worksheet.Range["B5"].Value2 = 2020;
            worksheet.Range["B6"].Value2 = 2021;
            worksheet.Range["B7"].Value2 = 2021;
            worksheet.Range["B8"].Value2 = 2021;
            worksheet.Range["B9"].Value2 = 2021;

            worksheet.Range["C1"].Value = "amount";
            worksheet.Range["C2"].Value2 = 50;
            worksheet.Range["C3"].Value2 = 60;
            worksheet.Range["C4"].Value2 = 70;
            worksheet.Range["C5"].Value2 = 80;
            worksheet.Range["C6"].Value2 = 90;
            worksheet.Range["C7"].Value2 = 100;
            worksheet.Range["C8"].Value2 = 110;
            worksheet.Range["C9"].Value2 = 120;

            // Get slicer collection
            XlsSlicerCollection slicers = worksheet.Slicers;

            //Create a table with the data from the specific cell range.
            IListObject table = worksheet.ListObjects.Create("Super Table", worksheet.Range["A1:C9"]);

            int count = 3;
            int index = 0;
            foreach (SlicerStyleType type in Enum.GetValues(typeof(SlicerStyleType)))
            {
                count += 5;
                String range = "E" + count;
                index = slicers.Add(table, range.ToString(), 0);

                //Style setting
                XlsSlicer xlsSlicer = slicers[index];
                xlsSlicer.Name = "slicers_" + count;
                xlsSlicer.StyleType = type;
            }

            //Save to file
            wb.SaveToFile("CreateSlicerFromTable.xlsx", ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            wb.Dispose();

            // Launch the file
            ExcelDocViewer("CreateSlicerFromTable.xlsx");
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
