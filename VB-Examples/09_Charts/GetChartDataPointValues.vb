Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Imports System.Text
Imports System.IO

Namespace GetChartDataPointValues
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a StringBuilder object to store the extracted data.
            Dim stringBuilder As New StringBuilder()

            ' Create a new Workbook object and load an Excel file.
            Dim workbook As New Workbook()
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToImage.xlsx")

            ' Get the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart in the worksheet.
            Dim chart As Chart = sheet.Charts(0)

            ' Get the first series in the chart.
            Dim cs As ChartSerie = chart.Series(0)

            ' Loop through each cell range in the series values.
            For Each cr As CellRange In cs.Values
                ' Append the range address to the StringBuilder.
                stringBuilder.Append(cr.RangeAddress & vbCrLf)

                ' Append the value of the data point to the StringBuilder.
                stringBuilder.Append("The value of the data point is " & cr.Value & vbCrLf)
            Next cr

            ' Specify the file name for the result.
            Dim result As String = "result.txt"

            ' Write the contents of the StringBuilder to a text file.
            File.WriteAllText(result, stringBuilder.ToString())
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
