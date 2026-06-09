using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace ExportEquations
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

            // Load an existing workbook with a pivot table from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExportEquations.xlsx");

            Worksheet sheet = workbook.Worksheets[0];

            // Export the first equation in the sheet to MathML format
            string mathML = sheet.Equations[0].ExportMathML();
            sheet.Range["B9"].Value = "mathML:";
            sheet.Range["B10"].Value = mathML;

            // Export the first equation to LaTeX format
            string LaTex = sheet.Equations[0].ExportLaTex();
            sheet.Range["B12"].Value = "LaTeX:";
            sheet.Range["B13"].Value = LaTex;

            // Specify the output file name for the result
            string outputFile = "ExportEquations_out.xlsx";
            string outputFile_TXT = "ExportEquations_TXT.txt";

            workbook.SaveToFile(outputFile);
            File.WriteAllText(outputFile_TXT, "LaTeX:\t" + LaTex + "\r\nmathML:\t" + mathML);

            // Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(outputFile, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
            FileViewer(outputFile);
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
