using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace NamedRanges
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }   
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\NamedRanges.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            //Creating a named range
            INamedRange NamedRange = workbook.NameRanges.Add("NewNamedRange");
            //Setting the range of the named range
            NamedRange.RefersToRange = sheet.Range["A8:E12"];

            String result = "NamedRanges_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);
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

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
