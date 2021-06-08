Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace ChangeDataLabel
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			'Create a workbook
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChangeDataLabel.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the chart
			Dim chart As Chart = sheet.Charts(0)

			'Change data label of the frist datapoint of the first series
			chart.Series(0).DataPoints(0).DataLabels.Text = "changed data label"

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
