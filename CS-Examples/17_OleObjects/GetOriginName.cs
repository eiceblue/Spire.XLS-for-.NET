using System;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;

namespace GetOriginName
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Load an existing workbook from a file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\GetOriginName.xlsx");

            // Get the first sheet
            Worksheet worksheet = workbook.Worksheets[0];

            String information = "";
            // Check if the worksheet contains any OLE objects
            if (worksheet.HasOleObjects)
            {
                // Iterate over each OLE object in the worksheet
                for (int i = 0; i < worksheet.OleObjects.Count; i++)
                {
                    // Get the current OLE object
                    var Object = worksheet.OleObjects[i];

                    // Determine the type of the OLE object
                    OleObjectType type = worksheet.OleObjects[i].ObjectType;
                    information += "Type: " + type.ToString()+"\n";

                    // Determine the origin name of the OLE object
                    String originName = worksheet.OleObjects[i].OleOriginName;
                    information += "Origin Name: " + originName + "\n";
                }
            }

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Save the information to a TXT file
            String result = "GetOriginName-out.txt";
            File.WriteAllText(result,information);

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
