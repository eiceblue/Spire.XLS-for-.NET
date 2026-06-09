using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace InsertBackgoundImageStream
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_1.xlsx");

            //Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Open the image as a stream
            Stream image = File.OpenRead(@"..\..\..\..\..\..\Data\Background.emf");

            //Set the image stream to be background image of the worksheet.
            sheet.PageSetup.BackgoundImageStream = image;

            // Specify the output file name for the EXCEL result
            string result = "InsertBackgoundImageStream_out.xlsx";

            //Save the Excel file
            workbook.SaveToFile(result, FileFormat.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

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
