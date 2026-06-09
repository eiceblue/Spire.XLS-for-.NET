Imports Spire.Xls
Imports System
Imports System.IO
Imports System.Windows.Forms

Namespace CreateExcelWithVBAMacro
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a workbook
            Dim workbook As Workbook = New Workbook()

            ' Add VBA project to the document
            Dim vbaProject As IVbaProject = workbook.VbaProject
            vbaProject.Name = "SampleVBAMacro"

            ' Record the original code page value
            Dim text As String = "Original code page: " + vbaProject.CodePage.ToString() + "\n"
            vbaProject.CodePage = 936; ' Set code page to 936 (Simplified Chinese support)

            text += "Modified code page: " + vbaProject.CodePage.ToString() + "\n"
            File.WriteAllText("CreateExcelWithVbaMacro.txt", text.ToString())

            ' Add a new VBA module to the project
            Dim vbaModule As IVbaModule = vbaProject.Modules.Add("SampleModule", VbaModuleType.Module)
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
                Dim ws As With
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
            End Sub"

            ' Save the workbook as Excel 97-2003 format (required for macro support)
            Dim result As String = "CreateExcelWithVbaMacro_out.xls"
            workbook.SaveToFile(result , FileFormat.Version97to2003)

            'View the document
            FileViewer(result)
        End Sub

        Private Sub FileViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

        Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
            Close()
        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class
End Namespace
