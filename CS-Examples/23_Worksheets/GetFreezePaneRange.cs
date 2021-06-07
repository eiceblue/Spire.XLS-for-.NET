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
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load an excel file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\GetFreezePaneRange.xlsx");

            Worksheet sheet = workbook.Worksheets[0];
            int rowIndex;
            int colIndex;

            //The row and column index of the frozen pane is passed through the out parameter. 
            //If it returns to 0, it means that it is not frozen
            sheet.GetFreezePanes(out rowIndex, out colIndex);

            string range = "Row index: " + rowIndex + ", column index: " + colIndex;

            //Save the document and launch it
            string result = "GetFreezePaneCellRange_result.txt";
            File.WriteAllText(result, range);
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
