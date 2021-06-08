Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace EachWorksheetToDifferentPDF
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load the document from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\EachWorksheetToDifferentPDFSample.xlsx")

			For Each sheet As Worksheet In workbook.Worksheets
				Dim FileName As String = sheet.Name & ".pdf"
				'Save the sheet to PDF
				sheet.SaveToPdf(FileName)

				'Launch the result file
				ExcelDocViewer(FileName)
			Next sheet

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
