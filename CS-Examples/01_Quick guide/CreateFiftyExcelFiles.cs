using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateFiftyExcelFiles
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
            for (int n = 0; n < 50; n++)
            {
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
                            sheet.Range[row, col].Text = "row" + row.ToString() + " col" + col.ToString();
                        }
                    }
                }

                workbook.SaveToFile("Workbook"+n+".xlsx", ExcelVersion.Version2010);
            }
            DateTime end = DateTime.Now;
            TimeSpan time = end - start;
            MessageBox.Show("50 File(s) have been created successfully! \n" + "Time consumed (Seconds): " + time.TotalSeconds.ToString());
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
