Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace FormatAxis

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the name of the worksheet to "FormatAxis"
            sheet.Name = "FormatAxis"

            ' Call the function to create chart data in the worksheet
            CreateChartData(sheet)

            ' Add a clustered column chart to the worksheet
            Dim chart As Chart = sheet.Charts.Add(ExcelChartType.ColumnClustered)

            ' Set the data range for the chart
            chart.DataRange = sheet.Range("B1:B9")

            ' Specify that the series data is not obtained from the range directly
            chart.SeriesDataFromRange = False

            ' Hide the plot area of the chart
            chart.PlotArea.Visible = False

            ' Set the positioning of the chart within the worksheet
            chart.TopRow = 10
            chart.BottomRow = 28
            chart.LeftColumn = 2
            chart.RightColumn = 10

            ' Set the chart title and customize its appearance
            chart.ChartTitle = "Chart with Customized Axis"
            chart.ChartTitleArea.IsBold = True
            chart.ChartTitleArea.Size = 12

            ' Get the first series of the chart
            Dim cs1 As Spire.Xls.Charts.ChartSerie = chart.Series(0)

            ' Set the category labels (X-axis values) for the first series
            cs1.CategoryLabels = sheet.Range("A2:A9")

            ' Customize the primary value axis of the chart
            chart.PrimaryValueAxis.MajorUnit = 8
            chart.PrimaryValueAxis.MinorUnit = 2
            chart.PrimaryValueAxis.MaxValue = 50
            chart.PrimaryValueAxis.MinValue = 0
            chart.PrimaryValueAxis.IsReverseOrder = False
            chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkOutside
            chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkInside
            chart.PrimaryValueAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionNextToAxis
            chart.PrimaryValueAxis.CrossesAt = 0

            ' Set the number format for the primary value axis to show currency format
            chart.PrimaryValueAxis.NumberFormat = "$#,##0"

            ' Disable the link to the source for the primary value axis
            chart.PrimaryValueAxis.IsSourceLinked = False

            ' Get the first series of the chart again
            Dim serie As ChartSerie = chart.Series(0)

            ' Customize the data points of the series
            For Each dataPoint As ChartDataPoint In serie.DataPoints
                ' Set the fill type and foreground color of the data point
                dataPoint.DataFormat.Fill.FillType = ShapeFillType.SolidColor
                dataPoint.DataFormat.Fill.ForeColor = Color.LightGreen

                ' Set the transparency of the data point fill
                dataPoint.DataFormat.Fill.Transparency = 0.3
            Next dataPoint

            ' Save the modified workbook to a new file named "Output.xlsx" using Excel 2010 version
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Output.xlsx")
		End Sub
		Private Sub CreateChartData(ByVal sheet As Worksheet)

            ' Set the value of cell A1 to "Month".
            sheet.Range("A1").Value = "Month"

            ' Set the value of cell A2 to "Jan".
            sheet.Range("A2").Value = "Jan"

            ' Set the value of cell A3 to "Feb".
            sheet.Range("A3").Value = "Feb"

            ' Set the value of cell A4 to "Mar".
            sheet.Range("A4").Value = "Mar"

            ' Set the value of cell A5 to "Apr".
            sheet.Range("A5").Value = "Apr"

            ' Set the value of cell A6 to "May".
            sheet.Range("A6").Value = "May"

            ' Set the value of cell A7 to "Jun".
            sheet.Range("A7").Value = "Jun"

            ' Set the value of cell A8 to "Jul".
            sheet.Range("A8").Value = "Jul"

            ' Set the value of cell A9 to "Aug".
            sheet.Range("A9").Value = "Aug"

            ' Set the value of cell B1 to "Planned".
            sheet.Range("B1").Value = "Planned"

            ' Set the numeric value of cell B2 to 38.
            sheet.Range("B2").NumberValue = 38

            ' Set the numeric value of cell B3 to 47.
            sheet.Range("B3").NumberValue = 47

            ' Set the numeric value of cell B4 to 39.
            sheet.Range("B4").NumberValue = 39

            ' Set the numeric value of cell B5 to 36.
            sheet.Range("B5").NumberValue = 36

            ' Set the numeric value of cell B6 to 27.
            sheet.Range("B6").NumberValue = 27

            ' Set the numeric value of cell B7 to 25.
            sheet.Range("B7").NumberValue = 25

            ' Set the numeric value of cell B8 to 36.
            sheet.Range("B8").NumberValue = 36

            ' Set the numeric value of cell B9 to 48.
            sheet.Range("B9").NumberValue = 48
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
