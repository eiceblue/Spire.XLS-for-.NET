using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace DisableDTD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            string outputFile_E =  "Ex.txt";
            try
            {
                string outputFile = "DisableDTD.xlsx";
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Disable DTD
            workbook.ProhibitDtd = true;  

            //Load the file from disk.
             workbook.LoadFromFile(@"..\..\..\..\..\..\Data\haveDtd.xlsx");

            //Save
             workbook.SaveToFile(outputFile, ExcelVersion.Version2013);
       
            // Dispose of the workbook object
            workbook.Dispose();

             FileViewer(outputFile);
            }
            catch (Exception ex)
            {
                FileStream stream = new FileStream(outputFile_E, FileMode.Append);
                StreamWriter sw = new StreamWriter(stream);
                sw.WriteLine(ex + "Disable DTD processing：" + ex.ToString());
                sw.Flush();
                sw.Close();
                FileViewer(outputFile_E);
            }

            this.Close();

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
