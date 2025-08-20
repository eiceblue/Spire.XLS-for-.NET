using System;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;

namespace ToSVG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load a file from the specified path into the workbook
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToSVG.xlsx");

            // Iterate through each worksheet in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Create a FileStream to write the SVG content to a file
                FileStream fs = new FileStream(string.Format("sheet{0}.svg", i), FileMode.Create);
                // Convert the worksheet to SVG and write it to the FileStream
                workbook.Worksheets[i].ToSVGStream(fs, 0, 0, 0, 0);
                // Flush and close the FileStream to ensure data is written and resources are released
                fs.Flush();
                fs.Close();
            }

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the document
            FileViewer("sheet0.svg");
            
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
