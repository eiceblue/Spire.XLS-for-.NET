Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Charts

Namespace SetLegendBackgroundColor

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.xlsx")

			Dim ws As Worksheet = workbook.Worksheets(0)
			Dim chart As Chart = ws.Charts(0)

			Dim x As XlsChartFrameFormat = TryCast(chart.Legend.FrameFormat, XlsChartFrameFormat)
			x.Fill.FillType = ShapeFillType.SolidColor
			x.ForeGroundColor = Color.SkyBlue

			'Save the document
			Dim output As String = "SetLegendBackgroundColor.xlsx"
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
