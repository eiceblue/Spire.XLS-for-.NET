Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace CreateParetoChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "ParetoChart.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ParetoChart.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet = workbook.Worksheets(0)

            ' Add a chart to the worksheet
            Dim officeChart = sheet.Charts.Add()

            ' Set the chart type to Pareto
            officeChart.ChartType = ExcelChartType.Pareto

            ' Set the data range for the chart
            officeChart.DataRange = sheet("A2:B8")

            ' Set the position of the chart within the worksheet using row and column indices
            officeChart.TopRow = 1
            officeChart.BottomRow = 19
            officeChart.LeftColumn = 4
            officeChart.RightColumn = 12

            ' Enable binning on the primary category axis (X-axis)
            officeChart.PrimaryCategoryAxis.IsBinningByCategory = True

            ' Set the value to assign to overflow bins
            officeChart.PrimaryCategoryAxis.OverflowBinValue = 5

            ' Set the value to assign to underflow bins
            officeChart.PrimaryCategoryAxis.UnderflowBinValue = 1

            ' Set the color of the Pareto line to blue
            officeChart.Series(0).ParetoLineFormat.LineProperties.Color = Color.Blue

            ' Set the gap width between bars in the chart series
            officeChart.Series(0).DataFormat.Options.GapWidth = 6

            ' Set the chart title to "Expenses"
            officeChart.ChartTitle = "Expenses"

            ' Disable the legend in the chart
            officeChart.HasLegend = False

            ' Save the modified workbook to "Pareto_chart.xlsx"
            workbook.SaveToFile("Pareto_chart.xlsx")
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("Pareto_chart.xlsx")
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
