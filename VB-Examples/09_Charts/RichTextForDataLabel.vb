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
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToImage.xlsx")

			'Get first worksheet of the workbook
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Get the first chart inside this worksheet
			Dim chart As Chart = worksheet.Charts(0)

			'Get the first datalabel of the first series 
			Dim datalabel As ChartDataLabels = chart.Series(0).DataPoints(0).DataLabels

			'Set the text
			datalabel.Text = "Rich Text Label"

			'Show the value
			chart.Series(0).DataPoints(0).DataLabels.HasValue = True

			'Set styles for the text
			chart.Series(0).DataPoints(0).DataLabels.Font.Color = Color.Red
			chart.Series(0).DataPoints(0).DataLabels.Font.IsBold = True

			'String for output file 
			Dim outputFile As String = "Output.xlsx"

			'Save the file
			workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

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
