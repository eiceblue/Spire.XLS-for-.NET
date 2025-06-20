Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.ComponentModel
Imports System.Text

Namespace DiscontinuousData
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim book As New Workbook()

            ' Load an existing Excel file named "DiscontinuousData.xlsx"
            book.LoadFromFile("..\..\..\..\..\..\Data\DiscontinuousData.xlsx")

            ' Get the first worksheet from the loaded workbook and assign it to a variable named "sheet"
            Dim sheet As Worksheet = book.Worksheets(0)

            ' Add a ColumnClustered chart to the worksheet and assign it to a variable named "chart"
            Dim chart As Chart = sheet.Charts.Add(ExcelChartType.ColumnClustered)

            ' Set the SeriesDataFromRange property to False, indicating that series data will be set explicitly
            chart.SeriesDataFromRange = False

            ' Specify the position and size of the chart
            chart.LeftColumn = 1
            chart.TopRow = 10
            chart.RightColumn = 10
            chart.BottomRow = 24

            ' Add a new ChartSerie to the chart and assign it to a variable named "cs1"
            Dim cs1 As ChartSerie = CType(chart.Series.Add(), ChartSerie)

            ' Set the name of the series based on the value in cell B1
            cs1.Name = sheet.Range("B1").Value

            ' Set the category labels for the series by combining ranges A2:A3, A5:A6, and A8:A9
            cs1.CategoryLabels = sheet.Range("A2:A3").AddCombinedRange(sheet.Range("A5:A6")).AddCombinedRange(sheet.Range("A8:A9"))

            ' Set the values for the series by combining ranges B2:B3, B5:B6, and B8:B9
            cs1.Values = sheet.Range("B2:B3").AddCombinedRange(sheet.Range("B5:B6")).AddCombinedRange(sheet.Range("B8:B9"))

            ' Set the type of the series as ColumnClustered
            cs1.SerieType = ExcelChartType.ColumnClustered

            ' Add another ChartSerie to the chart and assign it to a variable named "cs2"
            Dim cs2 As ChartSerie = CType(chart.Series.Add(), ChartSerie)

            ' Set the name of the series based on the value in cell C1
            cs2.Name = sheet.Range("C1").Value

            ' Set the category labels for the series by combining ranges A2:A3, A5:A6, and A8:A9
            cs2.CategoryLabels = sheet.Range("A2:A3").AddCombinedRange(sheet.Range("A5:A6")).AddCombinedRange(sheet.Range("A8:A9"))

            ' Set the values for the series by combining ranges C2:C3, C5:C6, and C8:C9
            cs2.Values = sheet.Range("C2:C3").AddCombinedRange(sheet.Range("C5:C6")).AddCombinedRange(sheet.Range("C8:C9"))

            ' Set the type of the series as ColumnClustered
            cs2.SerieType = ExcelChartType.ColumnClustered

            ' Set the title of the chart
            chart.ChartTitle = "Chart"

            ' Customize the appearance of the chart title
            chart.ChartTitleArea.Size = 20
            chart.ChartTitleArea.Color = Color.Black

            ' Disable major gridlines on the primary value axis of the chart
            chart.PrimaryValueAxis.HasMajorGridLines = False

            ' Save the workbook to a file named "Output.xlsx" in Excel 2010 format
            book.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            book.Dispose()
            ExcelDocViewer("Output.xlsx")
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
