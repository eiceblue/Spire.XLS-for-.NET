using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace AddTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a Workbook
            Workbook workbook = new Workbook();

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddTextBox.xlsx");
            
            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Add a Textbox
            ITextBoxLinkShape textbox = chart.Shapes.AddTextBox();

            // Set the width of the textbox
            textbox.Width = 1200;
            // Set the height of the textbox
            textbox.Height = 320;
            // Set the height of the textbox
            textbox.Left = 1000;
            // Set the top position of the textbox
            textbox.Top = 480;
            textbox.Text = "This is a textbox";

            // Save the file
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Output.xlsx");
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
