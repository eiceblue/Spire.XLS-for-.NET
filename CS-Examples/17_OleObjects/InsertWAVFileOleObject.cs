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
            //Create a workbook
            Workbook workbook = new Workbook();

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Add OLE object
            IOleObject oleObject = sheet.OleObjects.Add(@"..\..\..\..\..\..\Data\WAVFileSample.wav", Image.FromFile(@"..\..\..\..\..\..\Data\SpireXls.png"), OleLinkType.Embed);
            //Set the object location
            oleObject.Location = sheet.Range["B4"];
            //Set the object type as package
            oleObject.ObjectType = OleObjectType.Package;

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
