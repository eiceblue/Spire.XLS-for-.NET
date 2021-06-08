Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddSpinnerControl
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set text for range C11
			sheet.Range("C11").Text = "Value:"
			sheet.Range("C11").Style.Font.IsBold = True

			'Set value for range B10
			sheet.Range("C12").Value2 = 0

			'Add spinner control
			Dim spinner As ISpinnerShape = sheet.SpinnerShapes.AddSpinner(12, 4, 20, 20)
			spinner.LinkedCell = sheet.Range("C12")
			spinner.Min = 0
			spinner.Max = 100
			spinner.IncrementalChange = 5
			spinner.Display3DShading = True

			'Save the document
			Dim output As String = "AddSpinnerControl_out.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the Excel file
			ExcelDocViewer(output)
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
