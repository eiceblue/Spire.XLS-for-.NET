Imports Spire.Xls
Imports System
Imports System.Windows.Forms

Namespace MergeScenario
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a workbook
            Dim wb As Workbook = New Workbook()
            Dim inputFile As String = @"..\..\..\..\..\..\Data\ScenarioSample4.xlsx"
            wb.LoadFromFile(inputFile)
            Dim worksheet1 As Worksheet = wb.Worksheets[0]
            Dim worksheet2 As Worksheet = wb.Worksheets[1]

            'Merge the scenario 
            worksheet1.Scenarios.Merge(worksheet2)

            'Saving the workbook
            Dim outputFile As String = "MergeScenario.xlsx"
            wb.SaveToFile(outputFile, ExcelVersion.Version2013)
            wb.Dispose()

            'View the document
            FileViewer(outputFile)
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
