Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace LockSpecificColumnInNewExcel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Create an empty worksheet.
			workbook.CreateEmptySheet()

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Loop through all the columns in the worksheet and unlock them.
			For i As Integer = 0 To 254
				sheet.Rows(i).Style.Locked = False
			Next i

			'Lock the fourth column in the worksheet.
			sheet.Columns(3).Text = "Locked"
			sheet.Columns(3).Style.Locked = True

			'Set the password.
			sheet.Protect("123", SheetProtectionType.All)

			Dim result As String = "Result-LockSpecificColumnInNewlyXlsFile.xlsx"

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
