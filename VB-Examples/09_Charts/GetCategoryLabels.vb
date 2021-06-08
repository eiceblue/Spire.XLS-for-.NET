Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Imports System.Text
Imports System.IO

Namespace GetCategoryLabels
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim sb As New StringBuilder()

			'Create a workbook
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SampeB_4.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the chart
			Dim chart As Chart = sheet.Charts(0)

			'Get the cell range of the category labels
			Dim cr As CellRange = chart.PrimaryCategoryAxis.CategoryLabels
			For Each cell In cr
				sb.Append(cell.Value & vbCrLf)
			Next cell

			'Save and launch result file
			Dim result As String = "result.txt"
			File.WriteAllText(result, sb.ToString())
			ExcelDocViewer(result)

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
