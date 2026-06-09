using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace CreateExcelWithVbaMacro
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

            // Build the VBA project information string
            string text = "IsProtected：" + vbaProject.IsProtected + "\n";
            text += "Name：" + vbaProject.Name + "\n";
            text += "Description：" + vbaProject.Description + "\n";
            text += "LockProjectView：" + vbaProject.LockProjectView + "\n";
            text += "CodePage：" + vbaProject.CodePage + "\n";

            // Loop through all VBA modules (including standard modules and worksheet modules)
            foreach (IVbaModule module in vbaProject.Modules)
            {
                text += "\n--- Module ---\n";
                text += "Name：" + module.Name + "\n";
                text += "Type：" + module.Type.ToString() + "\n";
                text += "SourceCode：\n" + (string.IsNullOrEmpty(module.SourceCode) ? "(No SourceCode)" : module.SourceCode) + "\n";
            }

            File.WriteAllText("ExtracrVBAMacroProjectInfo.txt", text.ToString());
            
            // Clear all VBA modules from the project
            vbaProject.Modules.Clear();

            // Save the modified workbook (without macros) to a new file
            String result = "ExtracrVBAMacroProjectInfo.xls";
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
