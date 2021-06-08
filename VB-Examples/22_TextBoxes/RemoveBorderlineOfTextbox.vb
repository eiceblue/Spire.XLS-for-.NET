Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace RemoveBorderlineOfTextbox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()
			workbook.Version = ExcelVersion.Version2013

			'Create a new worksheet named "Remove Borderline" and add a chart to the worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)
			sheet.Name = "Remove Borderline"
			Dim chart As Chart = sheet.Charts.Add()

			'Create textbox1 in the chart and input text information.
			Dim textbox1 As XlsTextBoxShape = TryCast(chart.TextBoxes.AddTextBox(50, 50, 100, 600), XlsTextBoxShape)
			textbox1.Text = "The solution with borderline"

			'Create textbox2 in the chart, input text information and remove borderline.
			Dim textbox2 As XlsTextBoxShape = TryCast(chart.TextBoxes.AddTextBox(1000, 50, 100, 600), XlsTextBoxShape)
			textbox2.Text = "The solution without borderline"
			textbox2.Line.Weight = 0

			Dim result As String = "Result-RemoveBorderlineOfTextbox.xlsx"

			'Save to file.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the MS Excel file.
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
