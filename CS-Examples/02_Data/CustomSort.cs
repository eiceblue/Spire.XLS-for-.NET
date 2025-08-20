using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CustomSort
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create Excel document
            Workbook workbook = new Workbook();

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set header to participate in sorting
            workbook.DataSorter.IsIncludeTitle = false;
            // Add data
            sheet.Range["A1"].Text = "AA";
            sheet.Range["A2"].Text = "BB";
            sheet.Range["A3"].Text = "CC";
            sheet.Range["A4"].Text = "DD";
            sheet.Range["A5"].Text = "EE";
            sheet.Range["A6"].Text = "FF";
            sheet.Range["A7"].Text = "GG";
            sheet.Range["A8"].Text = "HH";
            // Custom sort
            workbook.DataSorter.SortColumns.Add(0, new String[]
                {"DD","CC", "BB", "AA", "HH","GG","FF","EE"});
            workbook.DataSorter.Sort(workbook.Worksheets[0].Range["A1:A8"]);

            // Specify the name for the resulting Excel file
            String result = "result.xlsx";
            // Save the document
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to free up resources
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
