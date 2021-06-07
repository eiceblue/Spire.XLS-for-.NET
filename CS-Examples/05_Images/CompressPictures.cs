using Spire.Xls;
using System;
using System.Windows.Forms;

namespace CompressPictures
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

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CompressPictures.xlsx");

            //Get the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (ExcelPicture picture in sheet.Pictures)
                {
                    picture.Compress(50);
                }
            }

            //Save to file
            String result = "CompressPictures_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

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
    }
}
