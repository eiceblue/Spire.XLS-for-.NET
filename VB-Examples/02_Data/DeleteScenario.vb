Imports Spire.Xls
Imports System
Imports System.Globalization
Imports System.IO
Imports System.Threading
Imports System.Windows.Forms

Namespace DeleteScenario
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a workbook
            Dim wb As Workbook = New Workbook()
            Dim inputFile As String = @"..\..\..\..\..\..\Data\ScenarioSample2.xlsx"
            wb.LoadFromFile(inputFile)
            Dim worksheet As Worksheet = wb.Worksheets[0]

            ' Access the collection of scenarios in the worksheet
            Dim scenarios As XlsScenarioCollection = worksheet.Scenarios

            'delete the scenario 
            scenarios.RemoveScenarioAt(0)
            scenarios.RemoveScenarioByName("two")

            Dim content As String = ""
            content += "Count:" + scenarios.Count + "\n"
            content += "ContainsScenario:" + scenarios.ContainsScenario("two").ToString() + "\n"
            content += "ContainsScenario:" + scenarios.ContainsScenario("one").ToString() + "\n"
            content += "ContainsScenario:" + scenarios.ContainsScenario("three").ToString() + "\n"
            File.WriteAllText("DeleteScenario.txt", content.ToString())

            'Saving the workbook
            Dim outputFile As String = "DeleteScenario.xlsx"
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
