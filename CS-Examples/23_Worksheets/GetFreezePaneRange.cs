using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace GetFreezePaneRange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\GetFreezePaneRange.xlsx");

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];
            int rowIndex;
            int colIndex;

            //The row and column index of the frozen pane is passed through the out parameter. 
            //If it returns to 0, it means that it is not frozen
            sheet.GetFreezePanes(out rowIndex, out colIndex);

            string range = "Row index: " + rowIndex + ", column index: " + colIndex;

            // Specify the output file path and name
            string result = "GetFreezePaneCellRange_result.txt";

            // Save the file
            File.WriteAllText(result, range);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            FileViewer(result);
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
