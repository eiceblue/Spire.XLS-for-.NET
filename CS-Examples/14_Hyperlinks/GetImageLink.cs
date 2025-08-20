using System;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;

namespace GetImageLink
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\hyperlink.xlsx");

            //Get the first picture of the first worksheet
            ExcelPicture picture = workbook.Worksheets[0].Pictures[0];

            //Get the address
            string address = picture.GetHyperLink().Address;

            // Write the address to the txt file
            string file = "address.txt";
            File.WriteAllText(file, address);

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
