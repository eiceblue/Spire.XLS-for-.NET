using Spire.Xls;
using System;
using System.Windows.Forms;

namespace RemoveVBAMacroProject
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook wb = new Workbook();

            //Load the file from disk.
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelWithVbaMacroProject.xls");

            //Get the first worksheet.
            Worksheet ws = wb.Worksheets[0];
            IVbaProject vbaProject = wb.VbaProject;

            // Remove a specific module by its name
            vbaProject.Modules.Remove("SampleModule");

            // Remove a module at the specified index 
            //vbaProject.Modules.RemoveAt(0);

            // Save the modified workbook (without macros) to a new file
            String result = "RemoveVBAMacroProject.xls";
            wb.SaveToFile(result);

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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
