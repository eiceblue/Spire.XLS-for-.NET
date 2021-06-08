Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace FitWidthWhenConvertToPDF
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load the document from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")

			For Each sheet As Worksheet In workbook.Worksheets
				'Auto fit page height
				sheet.PageSetup.FitToPagesTall = 0
				'Fit one page width
				sheet.PageSetup.FitToPagesWide = 1
			Next sheet

			'Save and launch result file
			Dim result As String = "result.pdf"
			workbook.SaveToFile(result, FileFormat.PDF)
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
