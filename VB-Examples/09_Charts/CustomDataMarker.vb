Imports Spire.Xls

Namespace CustomDataMarker

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			workbook.CreateEmptySheets(1)
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Add some sample data
			sheet.Name = "Demo"
			sheet.Range("A1").Value = "Tom"
			sheet.Range("A2").NumberValue = 1.5
			sheet.Range("A3").NumberValue = 2.1
			sheet.Range("A4").NumberValue = 3.6
			sheet.Range("A5").NumberValue = 5.2
			sheet.Range("A6").NumberValue = 7.3
			sheet.Range("A7").NumberValue = 3.1
			sheet.Range("B1").Value = "Kitty"
			sheet.Range("B2").NumberValue = 2.5
			sheet.Range("B3").NumberValue = 4.2
			sheet.Range("B4").NumberValue = 1.3
			sheet.Range("B5").NumberValue = 3.2
			sheet.Range("B6").NumberValue = 6.2
			sheet.Range("B7").NumberValue = 4.7

			'Create a Scatter-Markers chart based on the sample data
			Dim chart As Chart = sheet.Charts.Add(ExcelChartType.ScatterMarkers)
			chart.DataRange = sheet.Range("A1:B7")
			chart.PlotArea.Visible = False
			chart.SeriesDataFromRange = False
			chart.TopRow = 5
			chart.BottomRow = 22
			chart.LeftColumn = 4
			chart.RightColumn = 11
			chart.ChartTitle = "Chart with Markers"
			chart.ChartTitleArea.IsBold = True
			chart.ChartTitleArea.Size = 10

			'Format the markers in the chart by setting the background color, foreground color, type, size and transparency
			Dim cs1 As Spire.Xls.Charts.ChartSerie = chart.Series(0)
			cs1.DataFormat.MarkerBackgroundColor = Color.RoyalBlue
			cs1.DataFormat.MarkerForegroundColor = Color.WhiteSmoke
			cs1.DataFormat.MarkerSize = 7
			cs1.DataFormat.MarkerStyle = ChartMarkerType.PlusSign
			cs1.DataFormat.MarkerTransparencyValue = 0.8

			Dim cs2 As Spire.Xls.Charts.ChartSerie = chart.Series(1)
			cs2.DataFormat.MarkerBackgroundColor = Color.Pink
			cs2.DataFormat.MarkerSize = 9
			cs2.DataFormat.MarkerStyle = ChartMarkerType.Triangle
			cs2.DataFormat.MarkerTransparencyValue = 0.9


			'Save the document
			Dim output As String = "CustomDataMarker.xlsx"
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
