using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace AddLabelControl
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

            // Load the Excel document from the specified file path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Add a label control to the worksheet
            ILabelShape label = sheet.LabelShapes.AddLabel(10, 2, 30, 200);

            // Set the text content of the label control
            label.Text = "This is a Label Control";

            //Specify the filename for the resulting Excel file
            string output = "InsertLabelControl_out.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
            ExcelDocViewer(output);
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
