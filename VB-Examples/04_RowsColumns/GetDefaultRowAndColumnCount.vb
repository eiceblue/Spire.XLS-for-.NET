Imports System.IO
Imports System.Text
Imports Spire.Xls

Namespace GetDefaultRowAndColumnCount

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			'Clear all worksheets
			workbook.Worksheets.Clear()

			'Create a new worksheet
			Dim sheet As Worksheet = workbook.CreateEmptySheet()
			Dim sb As New StringBuilder()
			'Get row and column count
			Dim rowCount As Integer = sheet.Rows.Length
			Dim columnCount As Integer = sheet.Columns.Length

			sb.AppendLine("The default row count is :" & rowCount)
			sb.AppendLine("The default column count is :" & columnCount)

			'Save to Text file
			Dim output As String = "GetDefaultRowAndColumnCount.txt"
			File.WriteAllText(output, sb.ToString())

			'Launch the file
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
