Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace GetWorksheetOfChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToImage.xlsx")

			'Access first worksheet of the workbook
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Access the first chart inside this worksheet
			Dim chart As Chart = worksheet.Charts(0)

			'Get its worksheet
			Dim wSheet As Worksheet = TryCast(chart.Worksheet, Worksheet)

			'Create StringBuilder to save 
			Dim content As New StringBuilder()

			'Set string format for displaying
			Dim result As String = String.Format("Sheet Name: " & worksheet.Name & vbCrLf & "Charts' sheet Name: " & wSheet.Name)

			'Add result string to StringBuilder
			content.AppendLine(result)

			'String for output file 
			Dim outputFile As String = "Output.txt"

			'Save them to a txt file
			File.WriteAllText(outputFile, content.ToString())

			'Launching the output file.
			Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
