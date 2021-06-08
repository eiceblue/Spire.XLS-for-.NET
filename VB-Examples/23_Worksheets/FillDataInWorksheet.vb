Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace FillDataInWorksheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Get first worksheet of the workbook
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Fill data
			worksheet.Range("A1").Style.Font.IsBold = True
			worksheet.Range("B1").Style.Font.IsBold = True
			worksheet.Range("C1").Style.Font.IsBold = True
			worksheet.Range("A1").Text = "Month"
			worksheet.Range("A2").Text = "January"
			worksheet.Range("A3").Text = "February"
			worksheet.Range("A4").Text = "March"
			worksheet.Range("A5").Text = "April"
			worksheet.Range("B1").Text = "Payments"
			worksheet.Range("B2").NumberValue = 251
			worksheet.Range("B3").NumberValue = 515
			worksheet.Range("B4").NumberValue = 454
			worksheet.Range("B5").NumberValue = 874
			worksheet.Range("C1").Text = "Sample"
			worksheet.Range("C2").Text = "Sample1"
			worksheet.Range("C3").Text = "Sample2"
			worksheet.Range("C4").Text = "Sample3"
			worksheet.Range("C5").Text = "Sample4"

			'Set width for the second column
			worksheet.SetColumnWidth(2, 10)

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
