Imports Spire.Xls
Imports System
Imports System.Windows.Forms

Namespace EditScenario
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a workbook
            Dim wb As Workbook = New Workbook()
            Dim inputFile As String = @"..\..\..\..\..\..\Data\ScenarioSample3.xlsx"
            wb.LoadFromFile(inputFile)
            Dim worksheet As Worksheet = wb.Worksheets[0]

            ' Access the collection of scenarios in the worksheet
            Dim scenarios As XlsScenarioCollection = worksheet.Scenarios
            Dim scenario1 As XlsScenario = scenarios[0]
            Dim scenario2 As XlsScenario = scenarios[1]

            'Modify the scenario 
            scenario1.SetVariableCells(worksheet.Range["A2:A7"], scenario2.Values)

            Dim sourceCell As CellRange = worksheet.Range["B2:B7"]
            scenario2.SetVariableCells(sourceCell, scenario2.Values)

            scenario1.Show()
            scenario2.Show()

            'Saving the workbook
            Dim outputFile As String = "EditScenario.xlsx"
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
