Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Imports System.Text
Imports System.IO

Namespace EditLineChart
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\LineChart.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the line chart
			Dim chart As Chart = sheet.Charts(0)

			'Add a new series
			Dim cs As ChartSerie = chart.Series.Add("Added")

			'Set the values for the series
			cs.Values = sheet.Range("I1:L1")

			'Save and launch result file
			Dim result As String = "result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)
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
