Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Imports System.Text
Imports System.IO

Namespace GetChartDataPointValues
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim sb As New StringBuilder()

			'Load the document from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToImage.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the chart
			Dim chart As Chart = sheet.Charts(0)

			'Get the first series of the chart
			Dim cs As ChartSerie = chart.Series(0)

			For Each cr As CellRange In cs.Values
				sb.Append(cr.RangeAddress & vbCrLf)

				'Get the data point value
				sb.Append("The value of the data point is " & cr.Value & vbCrLf)
			Next cr

			Dim result As String = "result.txt"
			'Save and launch result file
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
