Imports Spire.Xls
Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms

Namespace CreateScenario
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a workbook
            Dim wb As Workbook = New Workbook()
            Dim inputFile As String = @"..\..\..\..\..\..\Data\ScenarioSample1.xlsx"
            wb.LoadFromFile(inputFile)
            Dim worksheet As Worksheet = wb.Worksheets[0]

            ' Access the collection of scenarios in the worksheet
            Dim scenarios As XlsScenarioCollection = worksheet.Scenarios

            'Initialize list objects with different values for scenarios
            Dim currentChangePercentage_Values As List<Object> = New List<Object> { 0.23, 0.8, 1.1, 0.5, 0.35, 0.2 }
            Dim increasedChangePercentage_Values As List<Object> = New List<Object> { 0.45, 0.56, 0.9, 0.5, 0.58, 0.43 }
            Dim decreasedChangePercentage_Values As List<Object> = New List<Object> { 0.3, 0.2, 0.5, 0.3, 0.5, 0.23 }
            Dim currentQuantity_Values As List<Object> = New List<Object> { 1500, 3000, 5000, 4000, 500, 4000 }
            Dim increasedQuantity_Values As List<Object> = New List<Object> { 1000, 5000, 4500, 3900, 10000, 8900 }
            Dim decreasedQuantity_Values As List<Object> = New List<Object> { 1000, 2000, 3000, 3000, 300, 4000 }

            'Add scenarios in the worksheet with different values for the same cells
            scenarios.Add("Current % of Change", worksheet.Range["E2:E7"], currentChangePercentage_Values)
            scenarios.Add("Increased % of Change", worksheet.Range["E2:E7"], increasedChangePercentage_Values)
            scenarios.Add("Decreased % of Change", worksheet.Range["E2:E7"], decreasedChangePercentage_Values)
            scenarios.Add("Current Quantity", worksheet.Range["D2:D7"], currentQuantity_Values)
            scenarios.Add("Increased Quantity", worksheet.Range["D2:D7"], increasedQuantity_Values)
            scenarios.Add("Decreased Quantity", worksheet.Range["D2:D7"], decreasedQuantity_Values)

            'Saving the workbook
            Dim outputFile As String = "CreateScenario.xlsx"
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
