using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.IO;
using Spire.Xls.Core;

namespace InsertWavFileOLEObject
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Add an OLE object
            IOleObject oleObject = sheet.OleObjects.Add(@"..\..\..\..\..\..\Data\WAVFileSample.wav", Image.FromFile(@"..\..\..\..\..\..\Data\SpireXls.png"), OleLinkType.Embed);

            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            FileStream fs = new FileStream(@"..\..\..\..\..\..\Data\SpireXls.png", FileMode.Open, FileAccess.Read, FileShare.Read);
            byte[] bytes = new byte[fs.Length];
            fs.Read(bytes, 0, bytes.Length);
            fs.Close();
            Stream ImgFile = new MemoryStream(bytes);
            IOleObject oleObject = sheet.OleObjects.Add(@"..\..\..\..\..\..\Data\WAVFileSample.wav", ImgFile, OleLinkType.Embed);
            */
             
            // Set the location for the OLE object
            oleObject.Location = sheet.Range["B4"];

            // Set the type of the OLE object as a package
            oleObject.ObjectType = OleObjectType.Package;

            // Specify the output file name for the result
            string result = "result.xlsx";

            // Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
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
