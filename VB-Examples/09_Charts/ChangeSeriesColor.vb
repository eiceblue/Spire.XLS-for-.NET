Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.ComponentModel
Imports System.Text

Namespace ChangeSeriesColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "ChangeSeriesColor.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChangeSeriesColor.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = sheet.Charts(0)

            ' Get the second series from the chart
            Dim cs As ChartSerie = chart.Series(1)

            ' Set the fill type of the series to solid color
            cs.Format.Fill.FillType = ShapeFillType.SolidColor

            ' Set the foreground color of the series to orange
            cs.Format.Fill.ForeColor = Color.Orange

            ' Save the modified workbook to the specified file path as "Output.xlsx", using the Excel 2010 format
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Output.xlsx")
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
