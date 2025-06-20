Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Imports System.Text
Imports System.IO

Namespace EditLineChart
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "LineChart.xlsx"
            workbook.LoadFromFile("..\..\..\..\..\..\Data\LineChart.xlsx")

            ' Get the first worksheet from the loaded workbook and assign it to a variable named "sheet"
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet and assign it to a variable named "chart"
            Dim chart As Chart = sheet.Charts(0)

            ' Add a new ChartSerie with the name "Added" to the chart and assign it to a variable named "cs"
            Dim cs As ChartSerie = chart.Series.Add("Added")

            ' Set the values of the newly added series using the range I1:L1 from the worksheet
            cs.Values = sheet.Range("I1:L1")

            ' Specify the output file name as "result.xlsx"
            Dim result As String = "result.xlsx"

            ' Save the modified workbook to a file named "result.xlsx" in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer(result)
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
