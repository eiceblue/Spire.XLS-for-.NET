Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CopyVisibleSheets
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load a csv file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CopyVisibleSheets.xlsx")

			'Create a new workbook
			Dim workbookNew As New Workbook()
			workbookNew.Version = ExcelVersion.Version2013
			workbookNew.Worksheets.Clear()

			'Loop through the worksheets
			For Each sheet As Worksheet In workbook.Worksheets
				'Judge if the worksheet is visible
				If sheet.Visibility = WorksheetVisibility.Visible Then
					'Copy the sheet to new workbook
					Dim name As String = sheet.Name
					workbookNew.Worksheets.AddCopy(sheet)
				End If
			Next sheet

			'Save the Excel file
			Dim result As String = "CopyVisibleSheets_out.xlsx"
			workbookNew.SaveToFile(result, ExcelVersion.Version2013)
			ExcelDocViewer(result)
		End Sub

		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
		Private Sub btnClose_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
