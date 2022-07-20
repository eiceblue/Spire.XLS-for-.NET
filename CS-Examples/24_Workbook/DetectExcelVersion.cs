using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.Text;
using System.IO;

namespace DetectExcelVersion
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Files
            string[] files = new string[] { @"..\..\..\..\..\..\Data\ExcelSample97_N.xls", @"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx", @"..\..\..\..\..\..\Data\ExcelSample_N.xlsb" };

            StringBuilder builder = new StringBuilder();

            foreach (string file in files)
            {
                //Create a workbook
                Workbook workbook = new Workbook();

                //Load the document
                workbook.LoadFromFile(file);

                //Get the version
                ExcelVersion version = workbook.Version;

                builder.AppendLine(version.ToString());
            }

            //Save to txt file
            string result = "DetectExcelVersion_out.txt";
            File.WriteAllText(result, builder.ToString());

            //Launch the file
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
