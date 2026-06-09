Imports Spire.Xls
Imports System
Imports System.Windows.Forms

Namespace RemoveVBAMacroProject
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

            ' Remove a specific module by its name
            vbaProject.Modules.Remove("SampleModule")

            ' Remove a module at the specified index 
            'vbaProject.Modules.RemoveAt(0);

            ' Save the modified workbook (without macros) to a new file
            Dim result As String = "RemoveVBAMacroProject.xls"
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
