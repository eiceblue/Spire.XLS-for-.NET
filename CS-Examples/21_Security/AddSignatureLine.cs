using System;
using System.Windows.Forms;
using Spire.Xls;

namespace AddSignatureLine
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a workbook instance
            Workbook workbook = new Workbook();

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Add a signature line 
            sheet.Range["A1"].AddSignatureLine("Rose","manager", "manager@test.com", "a short text" ,false,true);

            //Save the file
            string file = "AddSignatureLine.xlsx";
            workbook.SaveToFile(file,ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            OutputViewer(file);
          
        }
        private void OutputViewer(string fileName)
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
