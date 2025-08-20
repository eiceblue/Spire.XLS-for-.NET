using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Shapes;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace DrawOneLineThroughTwoPoints
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //1)Draw a line according to relative position
            XlsLineShape line1 = worksheet.TypedLines.AddLine() as XlsLineShape;
            line1.LeftColumn = 3;
            line1.TopRow = 3;
            line1.LeftColumnOffset = 0;
            line1.TopRowOffset = 0; 

            line1.RightColumn = 4; 
            line1.BottomRow = 5; 
            line1.RightColumnOffset = 0;
            line1.BottomRowOffset = 0; 

            //2)Draw a line according to absolute position(pixels).
            XlsLineShape line2 = worksheet.TypedLines.AddLine() as XlsLineShape;
            line2.StartPoint = new Point(30, 50);
            line2.EndPoint = new Point(20, 80);

            //Save to file
            String result = "DrawOneLineThroughTwoPoints.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
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
    }
}
