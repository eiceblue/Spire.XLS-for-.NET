using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace CreateExcelWithVBAMacro
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

            // Add VBA project to the document
            IVbaProject vbaProject = workbook.VbaProject;
            vbaProject.Name = "SampleVBAMacro";

            // Record the original code page value
            string text = "Original code page: " + vbaProject.CodePage.ToString() + "\n";
            vbaProject.CodePage = 936; // Set code page to 936 (Simplified Chinese support)

            text += "Modified code page: " + vbaProject.CodePage.ToString() + "\n";
            File.WriteAllText("CreateExcelWithVbaMacro.txt", text.ToString());

            // Add a new VBA module to the project
            IVbaModule vbaModule = vbaProject.Modules.Add("SampleModule", VbaModuleType.Module);
            vbaModule.SourceCode = @"
            Sub ExampleMacro()
                 ' Declare variables
                Dim ws As Worksheet
                Dim i As Integer
                ' Set reference to the active worksheet
                Set ws = ActiveSheet 
                 ' Clear worksheet contents (optional)
                ws.Cells.Clear
                ' Populate sample data
                With ws
                    ' Write header row
                    .Range(""A1: C1"").Value = Array(""No."", ""Project Name"", ""Amount"")
                    ' Loop to populate 10 rows of data
                    For i = 1 To 10
                        .Cells(i + 1, 1).Value = i           ' Serial number column
                        .Cells(i + 1, 2).Value = ""Project "" & i   ' Project name column
                        .Cells(i + 1, 3).Value = i * 100     ' Amount column (sample calculation)
                    Next i
                     ' Auto-fit column widths
                    .Columns(""A:C"").AutoFit
                    ' Format header row
                    With.Range(""A1:C1"")
                        .Font.Bold = True
                        .Interior.Color = RGB(200, 220, 255) 
                    End With
                    ' Format amount column as currency
                    .Range(""C2:C11"").NumberFormat = ""$#,##0.00""
                End With
                ' Display completion message
                MsgBox ""Data population complete!"", vbInformation, ""Operation Prompt""
            End Sub";

            // Save the workbook as Excel 97-2003 format (required for macro support)
            String result = "CreateExcelWithVbaMacro_out.xls";
            workbook.SaveToFile(result , FileFormat.Version97to2003);

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
