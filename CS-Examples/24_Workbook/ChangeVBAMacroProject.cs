using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Vba;
using System;
using System.IO;
using System.Windows.Forms;

namespace ChangeVBAMacroProject
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
            Workbook wb = new Workbook();

            //Load the file from disk.
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelWithVbaMacroProject.xls");

            //Get the first worksheet.
            Worksheet ws = wb.Worksheets[0];
            IVbaProject vbaProject = wb.VbaProject;

            vbaProject.Password = "1234";
            vbaProject.Name = "modify";
            vbaProject.Description = "Description";
            vbaProject.HelpFileName = "image1.png";
            vbaProject.ConditionalCompilation = "DEBUG = 2";
            vbaProject.LockProjectView = true;

            IVbaModule mod = vbaProject.Modules.GetWorksheetModule(ws);
            mod.Name = "IVbaModule";
            mod.SourceCode = "Dim lRow As Long";
            mod.Type = VbaModuleType.Module;

            // Save the modified workbook (without macros) to a new file
            String result = "ChangeVBAMacroProject.xls";
            wb.SaveToFile(result);

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
