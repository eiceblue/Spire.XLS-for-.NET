Imports System.IO
Imports System.Text
Imports Spire.Xls
Imports Spire.Xls.Core

Namespace SetBorderWidthOfMarker

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SetBorderWidthOfMarker.xlsx")

			'Get the chart from the first worksheet
			Dim chart As Chart = workbook.Worksheets(0).Charts(0)

			chart.Series(0).DataFormat.MarkerBorderWidth = 1.5 'unit is pt

			chart.Series(1).DataFormat.MarkerBorderWidth = 2.5 'unit is pt


			Dim output As String = "SetBorderWidthOfMarker_out.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

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
