Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace GetWorksheetNames

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample3.xlsx")

			'Get the names of all worksheets
			Dim sb As New StringBuilder()
			For Each sheet As Worksheet In workbook.Worksheets
				sb.AppendLine(sheet.Name)
			Next sheet

			'Save to the Text file
			Dim output As String = "GetWorksheetNames.txt"
			File.WriteAllText(output, sb.ToString())

			'Launch the Excel file
			ExcelDocViewer(output)
		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
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
