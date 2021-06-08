Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace WrapOrUnwrapTextInCells
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Wrap the excel text;
			sheet.Range("C1").Text = "e-iceblue is in facebook and welcome to like us"
			sheet.Range("C1").Style.WrapText = True
			sheet.Range("D1").Text = "e-iceblue is in twitter and welcome to follow us"
			sheet.Range("D1").Style.WrapText = True

			'Unwrap the excel text;
			sheet.Range("C2").Text = "http://www.facebook.com/pages/e-iceblue/139657096082266"
			sheet.Range("C2").Style.WrapText = False
			sheet.Range("D2").Text = "https://twitter.com/eiceblue"
			sheet.Range("D2").Style.WrapText = False

			'Set the text color of Range["C1:D1"]
			sheet.Range("C1:D1").Style.Font.Size = 15
			sheet.Range("C1:D1").Style.Font.Color = Color.Blue
			'Set the text color of Range["C2:D2"]
			sheet.Range("C2:D2").Style.Font.Size = 15
			sheet.Range("C2:D2").Style.Font.Color = Color.DeepSkyBlue

			Dim result As String = "Result-WrapOrUnwrapTextInExcelCells.xlsx"

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
