using Spire.Xls;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace SetTheme
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
            Workbook srcWorkbook = new Workbook();
            // Load an excel file
            srcWorkbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetTheme.xlsx");

            // Get the first worksheet
            Worksheet srcWorksheet = srcWorkbook.Worksheets[0];
            
            // Create an empty workbook
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();
            workbook.Worksheets.AddCopy(srcWorksheet);

            // 1. Copy the theme of the workbook
            //workbook.CopyTheme(srcWorkbook);

            // 2. Set a certain type of color of the default theme in the workbook
            workbook.SetThemeColor(ThemeColorType.Dk1, Color.SkyBlue);

            // Save to file
            String result = "SetTheme_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources 
            workbook.Dispose();
            srcWorkbook.Dispose();

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
    }
}
