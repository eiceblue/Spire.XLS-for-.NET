Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Vba
Imports System
Imports System.IO
Imports System.Windows.Forms

Namespace ChangeVBAMacroProject
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

            vbaProject.Password = "1234"
            vbaProject.Name = "modify"
            vbaProject.Description = "Description"
            vbaProject.HelpFileName = "image1.png"
            vbaProject.ConditionalCompilation = "DEBUG = 2"
            vbaProject.LockProjectView = True

            Dim mod As IVbaModule = vbaProject.Modules.GetWorksheetModule(ws)
            mod.Name = "IVbaModule"
            mod.SourceCode = "Dim lRow As Long"
            mod.Type = VbaModuleType.Module

            ' Save the modified workbook (without macros) to a new file
            Dim result As String = "ChangeVBAMacroProject.xls"
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
