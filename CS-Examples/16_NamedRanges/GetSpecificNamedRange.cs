using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;
using System.Text;
using System.IO;

namespace GetSpecificNamedRange
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            //Load the document from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AllNamedRanges.xlsx");

            //Get specific named range by index
            string name1 = workbook.NameRanges[1].Name;
            sb.Append("Get the specific named range " + name1 + " by index" + "\r\n");


            //Get specific named range by name
            string name2 = workbook.NameRanges["NameRange3"].Name;
            sb.Append("Get the specific named range " + name2 + " by name" + "\r\n");

            //Save and launch result file
            string result = "result.txt";
            File.WriteAllText(result, sb.ToString());
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
