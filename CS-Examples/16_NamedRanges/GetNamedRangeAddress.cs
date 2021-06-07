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


namespace GetNamedRangeAddress
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

            //Create a workbook and load the document from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AllNamedRanges.xlsx");

            //Get specific named range by index
            INamedRange NamedRange = workbook.NameRanges[0];

            //Get the address of the named range
            string address = NamedRange.RefersToRange.RangeAddress;

            sb.Append("The address of the named range " + NamedRange.Name + " is " + address);

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
