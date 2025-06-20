Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace GaugeChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook and assign it to the "sheet" variable
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the name of the worksheet to "Gauge Chart"
            sheet.Name = "Gauge Chart"

            ' Call the function to create chart data on the sheet
            CreateChartData(sheet)

            ' Add a doughnut chart to the sheet and assign it to the "chart" variable
            Dim chart As Chart = sheet.Charts.Add(ExcelChartType.Doughnut)

            ' Set the data range for the chart to cells A1 to A5
            chart.DataRange = sheet.Range("A1:A5")

            ' Disable automatic series data detection from the range
            chart.SeriesDataFromRange = False

            ' Enable the legend for the chart
            chart.HasLegend = True

            ' Set the position of the chart within the worksheet
            chart.LeftColumn = 2
            chart.TopRow = 7
            chart.RightColumn = 9
            chart.BottomRow = 25

            ' Get the first series in the chart and assign it to the "cs1" variable
            Dim cs1 As ChartSerie = CType(chart.Series("Value"), ChartSerie)

            ' Set the size of the doughnut hole to 60% of the chart's diameter
            cs1.Format.Options.DoughnutHoleSize = 60

            ' Set the starting angle for the first slice of the doughnut to 270 degrees (anti-clockwise)
            cs1.DataFormat.Options.FirstSliceAngle = 270

            ' Set the fill color of the first data point in the series to yellow
            cs1.DataPoints(0).DataFormat.Fill.ForeColor = Color.Yellow

            ' Set the fill color of the second data point in the series to pale violet red
            cs1.DataPoints(1).DataFormat.Fill.ForeColor = Color.PaleVioletRed

            ' Set the fill color of the third data point in the series to dark violet
            cs1.DataPoints(2).DataFormat.Fill.ForeColor = Color.DarkViolet

            ' Make the fourth data point in the series invisible by setting its fill visibility to false
            cs1.DataPoints(3).DataFormat.Fill.Visible = False

            ' Add a new series of type "Pointer" (pie chart) to the chart and assign it to the "cs2" variable
            Dim cs2 As ChartSerie = CType(chart.Series.Add("Pointer", ExcelChartType.Pie), ChartSerie)

            ' Set the values for the second series from cells D2 to D4
            cs2.Values = sheet.Range("D2:D4")

            ' Use the secondary axis for the second series
            cs2.UsePrimaryAxis = False

            ' Enable data labels for the first data point in the second series
            cs2.DataPoints(0).DataLabels.HasValue = True

            ' Set the starting angle for the first slice of the second series to 270 degrees (anti-clockwise)
            cs2.DataFormat.Options.FirstSliceAngle = 270

            ' Make the first data point in the second series invisible by setting its fill visibility to false
            cs2.DataPoints(0).DataFormat.Fill.Visible = False

            ' Set the fill color of the second data point in the second series to black
            cs2.DataPoints(1).DataFormat.Fill.FillType = ShapeFillType.SolidColor
            cs2.DataPoints(1).DataFormat.Fill.ForeColor = Color.Black

            ' Make the third data point in the second series invisible by setting its fill visibility to false
            cs2.DataPoints(2).DataFormat.Fill.Visible = False

            ' Save the workbook to a file named "Output.xlsx" in the Excel 2010 format
            workbook.SaveToFile("Output.xlsx", FileFormat.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Output.xlsx")
		End Sub

		Private Sub CreateChartData(ByVal sheet As Worksheet)
            ' Set the value of cell A1 to "Value"
            sheet.Range("A1").Value = "Value"

            ' Set the value of cell A2 to "30"
            sheet.Range("A2").Value = "30"

            ' Set the value of cell A3 to "60"
            sheet.Range("A3").Value = "60"

            ' Set the value of cell A4 to "90"
            sheet.Range("A4").Value = "90"

            ' Set the value of cell A5 to "180"
            sheet.Range("A5").Value = "180"

            ' Set the value of cell C2 to "value"
            sheet.Range("C2").Value = "value"

            ' Set the value of cell C3 to "pointer"
            sheet.Range("C3").Value = "pointer"

            ' Set the value of cell C4 to "End"
            sheet.Range("C4").Value = "End"

            ' Set the value of cell D2 to "10"
            sheet.Range("D2").Value = "10"

            ' Set the value of cell D3 to "1"
            sheet.Range("D3").Value = "1"

            ' Set the value of cell D4 to "189"
            sheet.Range("D4").Value = "189"
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
