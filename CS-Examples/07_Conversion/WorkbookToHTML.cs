using Spire.Xls;
using System;
using System.Windows.Forms;

namespace WorkbookToHTML
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
            Workbook workbook = new Workbook();

            // Load a file from the specified path into the workbook
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorkbookToHTML.xlsx");

            //Convert to html
            workbook.SaveToHtml("result.html");

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
            FileViewer("result.html");
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
