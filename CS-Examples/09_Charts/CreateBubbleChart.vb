Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CreateBubbleChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "CreateBubbleChart.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateBubbleChart.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a Bubble chart to the worksheet
            Dim chart As Chart = sheet.Charts.Add(ExcelChartType.Bubble)

            ' Set the chart title to "Bubble"
            chart.ChartTitle = "Bubble"

            ' Configure the chart title area formatting
            chart.ChartTitleArea.IsBold = True
            chart.ChartTitleArea.Size = 12

            ' Set the data range for the chart
            chart.DataRange = sheet.Range("A1:C5")

            ' Set the series data to be manually specified
            chart.SeriesDataFromRange = False

            ' Set the bubbles data for the first series of the chart
            chart.Series(0).Bubbles = sheet.Range("C2:C5")

            ' Set the position of the chart within the worksheet
            chart.LeftColumn = 7
            chart.TopRow = 6
            chart.RightColumn = 16
            chart.BottomRow = 29

            ' Save the modified workbook to "CreateBubbleChart.xlsx" using Excel 2010 format
            workbook.SaveToFile("CreateBubbleChart.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer("CreateBubbleChart.xlsx")
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
