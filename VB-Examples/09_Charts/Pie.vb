Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace Pie
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook and set its name to "Pie Chart"
            Dim sheet As Worksheet = workbook.Worksheets(0)
            sheet.Name = "Pie Chart"

            ' Add a pie chart or a 3D pie chart based on the state of checkBox1
            Dim chart As Chart = Nothing
            If checkBox1.Checked Then
                chart = sheet.Charts.Add(ExcelChartType.Pie3D)
            Else
                chart = sheet.Charts.Add(ExcelChartType.Pie)
            End If

            ' Create the data for the chart
            CreateChartData(sheet)

            ' Set the data range for the chart and disable automatic series data detection
            chart.DataRange = sheet.Range("B2:B5")
            chart.SeriesDataFromRange = False

            ' Set the position and size of the chart on the worksheet
            chart.LeftColumn = 1
            chart.TopRow = 6
            chart.RightColumn = 9
            chart.BottomRow = 25

            ' Set the title of the chart and customize its appearance
            chart.ChartTitle = "Sales by year"
            chart.ChartTitleArea.IsBold = True
            chart.ChartTitleArea.Size = 12

            ' Get the first series in the chart and assign category labels and values ranges
            Dim cs As ChartSerie = chart.Series(0)
            cs.CategoryLabels = sheet.Range("A2:A5")
            cs.Values = sheet.Range("B2:B5")
            cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True

            ' Hide the fill color of the plot area
            chart.PlotArea.Fill.Visible = False

            ' Save the workbook to a file named "Output.xlsx" in Excel 2010 format
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Output.xlsx")
		End Sub

		Private Sub CreateChartData(ByVal sheet As Worksheet)
            'Set the value of cells A1 to B5
            sheet.Range("A1").Value = "Year"
            sheet.Range("A2").Value = "2002"
            sheet.Range("A3").Value = "2003"
            sheet.Range("A4").Value = "2004"
            sheet.Range("A5").Value = "2005"

            sheet.Range("B1").Value = "Sales"
            sheet.Range("B2").NumberValue = 4000
            sheet.Range("B3").NumberValue = 6000
            sheet.Range("B4").NumberValue = 7000
            sheet.Range("B5").NumberValue = 8500

            ' Set the row height of cells A1 to B1
            sheet.Range("A1:B1").RowHeight = 15
            ' Set the background color of cells A1 to B1
            sheet.Range("A1:B1").Style.Color = Color.DarkGray
            ' Set the font color of cells A1 to B1
            sheet.Range("A1:B1").Style.Font.Color = Color.White
            ' Set the vertical alignment of cells A1 to B1
            sheet.Range("A1:B1").Style.VerticalAlignment = VerticalAlignType.Center
            ' Set the horizontal alignment of cells A1 to B1
            sheet.Range("A1:B1").Style.HorizontalAlignment = HorizontalAlignType.Center
            'Apply a number format for cells  B2:C5
            sheet.Range("B2:C5").Style.NumberFormat = """$""#,##0"
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
