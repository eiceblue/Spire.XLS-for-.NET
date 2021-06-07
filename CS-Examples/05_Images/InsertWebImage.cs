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
            //Create a workbook
            Workbook workbook = new Workbook();

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            string URL = "http://www.e-iceblue.com/downloads/demo/Logo.png";

            //Instantiate the web client object
            WebClient webClient = new WebClient();

            //Extract image data into memory stream
            MemoryStream objImage = new System.IO.MemoryStream(webClient.DownloadData(URL));

            Image image = Image.FromStream(objImage);

            //Add the image in the sheet
            sheet.Pictures.Add(3, 2, image);
            
            //Save and launch result file
            string result = "result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);
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
