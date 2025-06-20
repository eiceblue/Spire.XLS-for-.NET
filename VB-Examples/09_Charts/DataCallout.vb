Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace DataCallout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "DataCallout.xlsx"
            workbook.LoadFromFile("..\..\..\..\..\..\Data\DataCallout.xlsx")

            ' Get the first worksheet from the loaded workbook and assign it to a variable named "sheet"
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet and assign it to a variable named "chart"
            Dim chart As Chart = sheet.Charts(0)

            ' Iterate through each series in the chart
            For Each cs As ChartSerie In chart.Series
                ' Enable data labels for the default data point of the series
                cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True

                ' Enable wedge callout for the data labels of the default data point
                cs.DataPoints.DefaultDataPoint.DataLabels.HasWedgeCallout = True

                ' Show category name in the data labels of the default data point
                cs.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = True

                ' Show series name in the data labels of the default data point
                cs.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = True

                ' Show legend key in the data labels of the default data point
                cs.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = True
            Next cs

            ' Save the workbook to a file named "Output.xlsx" in Excel 2010 format
            workbook.SaveToFile("Output.xlsx", FileFormat.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer(workbook.FileName)
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
