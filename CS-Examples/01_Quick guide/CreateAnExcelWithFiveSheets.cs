using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateAnExcelWithFiveSheets
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
            workbook.CreateEmptySheets(5);
            for (int i = 0; i < 5; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Name = "Sheet" + i.ToString();
                for (int row = 1; row <= 150; row++)
                {
                    for (int col = 1; col <= 50; col++)
                    {
                        sheet.Range[row, col].Text="row" + row.ToString() + " col" + col.ToString();
                    }
                }
            }

            String result = "CreateAnExcelWithFiveSheets_result.xlsx";

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
