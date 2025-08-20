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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractOle2.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Check if the worksheet contains any OLE objects
            if (sheet.HasOleObjects)
            {
                // Iterate over each OLE object in the worksheet
                for (int i = 0; i < sheet.OleObjects.Count; i++)
                {
                    // Get the current OLE object
                    var Object = sheet.OleObjects[i];

                    // Determine the type of the OLE object
                    OleObjectType type = sheet.OleObjects[i].ObjectType;

                    // Perform operations based on the type of the OLE object
                    switch (type)
                    {
                        // Word document
                        case OleObjectType.WordDocument:
                            File.WriteAllBytes("Ole.docx", Object.OleData);
                            break;
                        // PowerPoint document
                        case OleObjectType.PowerPointSlide:
                            File.WriteAllBytes("Ole.pptx", Object.OleData);
                            break;
                        // PDF document
                        case OleObjectType.AdobeAcrobatDocument:
                            File.WriteAllBytes("Ole.pdf", Object.OleData);
                            break;
                    }
                }
            }

            // Dispose of the workbook object to release resources
            workbook.Dispose();
            
            MessageBox.Show("Completed!");
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
