using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace ExtractMsgOLEObject
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            string outputFile = "test.msg";

            // Create a new workbook object
            Workbook workbook = new Workbook();

            //Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Msg.xlsx");
      
            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            OleObjectType type;
            // Determine if there is an ole object in the sheet
            if (sheet.HasOleObjects)
            {
                for (int i = 0; i < sheet.OleObjects.Count; i++)
                {
                    var Object = sheet.OleObjects[i];
                    // Get the type of ole object
                    type = sheet.OleObjects[i].ObjectType;
                    switch (type)
                    {
                        // If the type of ole object is msg
                        case OleObjectType.Msg:
                            File.WriteAllBytes(outputFile, Object.OleData);
                            // View the document using a file viewer
                            FileViewer(outputFile);
                            break;
                    }
                }
            }

            // Dispose of the workbook object
            workbook.Dispose();

            this.Close();
         
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
