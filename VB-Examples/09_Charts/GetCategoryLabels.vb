Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Imports System.Text
Imports System.IO

Namespace GetCategoryLabels
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new StringBuilder object named "stringBuilder"
            Dim stringBuilder As New StringBuilder()

            ' Create a new Workbook object named "workbook"
            Dim workbook As New Workbook()

            ' Load a workbook file named "SampeB_4.xlsx" from the specified relative path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampeB_4.xlsx")

            ' Get the first worksheet in the workbook and assign it to the "sheet" variable
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the chart at index 0 from the worksheet and assign it to the "chart" variable
            Dim chart As Chart = sheet.Charts(0)

            ' Get the cell range for the category labels of the primary category axis of the chart and assign it to the "cr" variable
            Dim cr As CellRange = chart.PrimaryCategoryAxis.CategoryLabels

            ' Iterate through each cell in the cell range
            For Each cell In cr
                ' Append the value of the current cell to the StringBuilder object "stringBuilder" with a line break (vbCrLf)
                stringBuilder.Append(cell.Value & vbCrLf)
            Next cell

            ' Set the output file name to "result.txt"
            Dim result As String = "result.txt"

            ' Write the contents of the StringBuilder object "stringBuilder" to the output file as text
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
