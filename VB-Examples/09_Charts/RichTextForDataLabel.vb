Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace RichTextForDataLabel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToImage.xlsx")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = worksheet.Charts(0)

            ' Get the data labels of the first series and first data point in the chart
            Dim datalabel As ChartDataLabels = chart.Series(0).DataPoints(0).DataLabels

            ' Set the text of the data label to "Rich Text Label"
            datalabel.Text = "Rich Text Label"

            ' Enable data labels for the first data point in the first series
            chart.Series(0).DataPoints(0).DataLabels.HasValue = True

            ' Set the font color of the data label to red
            chart.Series(0).DataPoints(0).DataLabels.Color = Color.Red

            ' Set the font style of the data label to bold
            chart.Series(0).DataPoints(0).DataLabels.IsBold = True

            ' Specify the output file name
            Dim outputFile As String = "Output.xlsx"

            ' Save the modified workbook to the specified file with Excel 2013 format
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
            Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
