using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateAnExcelWithOneSheet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            DateTime start = DateTime.Now;
            //Create a workbook
            Workbook workbook = new Workbook();
            workbook.CreateEmptySheets(1);
            Worksheet sheet = workbook.Worksheets[0];

            for (int row = 1; row <= 10000; row++)
            {
                for (int col = 1; col <= 30; col++)
                {
                    sheet.Range[row, col].Text = row.ToString() + "," + col.ToString();
                }
            }
            String result = "CreateAnExcelWithOneSheet_result.xlsx";

            workbook.SaveToFile(result, ExcelVersion.Version2010);

            DateTime end = DateTime.Now;
            TimeSpan time = end - start;
            MessageBox.Show("File has been created successfully! \n" + "Time consumed (Seconds): " + time.TotalSeconds.ToString());

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
