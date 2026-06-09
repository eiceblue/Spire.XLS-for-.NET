Imports Spire.Xls
Imports System
Imports System.IO
Imports System.Windows.Forms

Namespace CreateExcelWithVbaMacro
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a workbook
            Dim wb As Workbook = New Workbook()

            'Load the file from disk.
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelWithVbaMacroProject.xls")

            'Get the first worksheet.
            Dim ws As Worksheet = wb.Worksheets[0]
            Dim vbaProject As IVbaProject = wb.VbaProject

            ' Build the VBA project information string
            Dim text As String = "IsProtected：" + vbaProject.IsProtected + "\n"
            text += "Name：" + vbaProject.Name + "\n"
            text += "Description：" + vbaProject.Description + "\n"
            text += "LockProjectView：" + vbaProject.LockProjectView + "\n"
            text += "CodePage：" + vbaProject.CodePage + "\n"

            ' Loop through all VBA modules (including standard modules and worksheet modules)
            For Each module As IVbaModule In vbaProject.Modules
                text += "\n -= 1- Module -= 1-\n"
                text += "Name：" + module.Name + "\n"
                text += "Type：" + module.Type.ToString() + "\n"
                text += "SourceCode：\n" + (String.IsNullOrEmpty(module.SourceCode) ? "(No SourceCode)" : module.SourceCode) + "\n"
            Next

            File.WriteAllText("ExtracrVBAMacroProjectInfo.txt", text.ToString())
            
            ' Clear all VBA modules from the project
            vbaProject.Modules.Clear()

            ' Save the modified workbook (without macros) to a new file
            Dim result As String = "ExtracrVBAMacroProjectInfo.xls"
            wb.SaveToFile(result)

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
