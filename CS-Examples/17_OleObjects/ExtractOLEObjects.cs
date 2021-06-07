using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.IO;

namespace ExtractOLEObjects
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractOle2.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Extract ole objects
            if (sheet.HasOleObjects)
            {
                for (int i = 0; i < sheet.OleObjects.Count; i++)
                {
                    var Object = sheet.OleObjects[i];
                    OleObjectType type = sheet.OleObjects[i].ObjectType;
                    switch (type)
                    {
                        //Word document
                        case OleObjectType.WordDocument:
                            File.WriteAllBytes("Ole.docx", Object.OleData);
                            break;
                        //PowerPoint document
                        case OleObjectType.PowerPointSlide:
                            File.WriteAllBytes("Ole.pptx", Object.OleData);
                            break;
                        //PDF document
                        case OleObjectType.AdobeAcrobatDocument:
                            File.WriteAllBytes("Ole.pdf", Object.OleData);
                            break;
                    }
                }
            }
            MessageBox.Show("Completed!");
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
