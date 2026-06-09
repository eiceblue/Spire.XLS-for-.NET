using Spire.Xls;
using System;
using System.Windows.Forms;

namespace LETAndMAPFunctions
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

            // Get the first sheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set number values for cells
            sheet.Range["A2"].NumberValue = 1;
            sheet.Range["A3"].NumberValue = 2;
            sheet.Range["A4"].NumberValue = 3;
            sheet.Range["B2"].NumberValue = 11;
            sheet.Range["B3"].NumberValue = 12;
            sheet.Range["B4"].NumberValue = 13;

            //Use the LET function
            sheet.Range["C1"].Text = "out";
            sheet.Range["C2"].Formula = "=LET(x, 5, y, 10, x + y)";
            sheet.Range["C3"].Formula = "=LET(a, 1, b, 2, c, 3, d, 4, a+b+c+d)";
            sheet.Range["C4"].Formula = "=LET(outer, LET(inner, 5, inner*2), outer+10)";

            //Use the MAP function
            sheet.Range["C2"].Formula = "=MAP(A2:A4, LAMBDA(x, x*2))";
            sheet.Range["D2"].Formula = "=MAP(A2:A4,LAMBDA(x,x*10+1))";
            sheet.Range["A8"].Formula = "=MAP(A2:B4,C2:D4,LAMBDA(x,y,SUM(x,y)))";

            // Recalculate all formulas to ensure values are up to date
            sheet.CalculateAllValue();

            // Save the modified workbook to the specified file using Excel 2010 format
            string result = @"LETAndMAPFunctions_out.xlsx";
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
