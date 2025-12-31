using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.IO;
using System.Net;

namespace InsertWebImage
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first sheet from the workbook.
            Worksheet sheet = workbook.Worksheets[0];

            // Specify the URL of the image to be downloaded.
            string URL = "http://www.e-iceblue.com/downloads/demo/Logo.png";

            // Instantiate a web client object.
            WebClient webClient = new WebClient();

            // Extract the image data into a memory stream.
            MemoryStream objImage = new System.IO.MemoryStream(webClient.DownloadData(URL));

            // Add the image to the worksheet at a specific location (row 3, column 2).
            sheet.Pictures.Add(3, 2, objImage);

            // Specify the resulting file name.
            string result = "result.xlsx";

            // Save the modified workbook to a file using Excel 2010 format.
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources.
            workbook.Dispose();

            // Launch the file
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
