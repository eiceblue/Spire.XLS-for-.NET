Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Imports System.Text
Imports System.IO

Namespace ChangeMajorGridlines
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			'Create a workbook
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SampeB_4.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the chart
			Dim chart As Chart = sheet.Charts(0)

			'Change the color of major gridlines
			chart.PrimaryValueAxis.MajorGridLines.LineProperties.Color = Color.Red

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
