Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace CreateBoxAndWhiskerChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "BoxAndWhiskerChart.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\BoxAndWhiskerChart.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet = workbook.Worksheets(0)

            ' Add a chart to the worksheet
            Dim officeChart = sheet.Charts.Add()

            ' Set the chart title to "Yearly Vehicle Sales"
            officeChart.ChartTitle = "Yearly Vehicle Sales"

            ' Set the chart type to Box and Whisker
            officeChart.ChartType = ExcelChartType.BoxAndWhisker

            ' Set the data range for the chart
            officeChart.DataRange = sheet("A1:E17")

            ' Get the first series of the chart
            Dim seriesA = officeChart.Series(0)

            ' Configure data format options for series A
            seriesA.DataFormat.ShowInnerPoints = False
            seriesA.DataFormat.ShowOutlierPoints = True
            seriesA.DataFormat.ShowMeanMarkers = True
            seriesA.DataFormat.ShowMeanLine = False
            seriesA.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian

            ' Get the second series of the chart
            Dim seriesB = officeChart.Series(1)

            ' Configure data format options for series B
            seriesB.DataFormat.ShowInnerPoints = False
            seriesB.DataFormat.ShowOutlierPoints = True
            seriesB.DataFormat.ShowMeanMarkers = True
            seriesB.DataFormat.ShowMeanLine = False
            seriesB.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.InclusiveMedian

            ' Get the third series of the chart
            Dim seriesC = officeChart.Series(2)

            ' Configure data format options for series C
            seriesC.DataFormat.ShowInnerPoints = False
            seriesC.DataFormat.ShowOutlierPoints = True
            seriesC.DataFormat.ShowMeanMarkers = True
            seriesC.DataFormat.ShowMeanLine = False
            seriesC.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian

            ' Save the workbook to a file named "Boxandwhisker_chart.xlsx"
            workbook.SaveToFile("Boxandwhisker_chart.xlsx")
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("Boxandwhisker_chart.xlsx")
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
