Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Charts

Namespace SetBorderColorAndStyle

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample3.xlsx")

			'Get the first worksheet from workbook and then get the first chart from the worksheet
			Dim ws As Worksheet = workbook.Worksheets(0)
			Dim chart As Chart = ws.Charts(0)

			'Set CustomLineWeight property for Series line
			TryCast(chart.Series(0).DataPoints(0).DataFormat.LineProperties, XlsChartBorder).CustomLineWeight = 2.5f
			'Set color property for Series line
			TryCast(chart.Series(0).DataPoints(0).DataFormat.LineProperties, XlsChartBorder).Color = Color.Red

			'Save the document
			Dim output As String = "SetBorderColorAndStyle.xlsx"
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
