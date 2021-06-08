Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace DifferentHeaderFooterOnFirstPage
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
			sheet.Range("A1").Text="Hello World"
			sheet.Range("F30").Text = "Hello World"
			sheet.Range("G150").Text = "Hello World"

			'Set the value to show the headers/footers for first page are different from the other pages.
			sheet.PageSetup.DifferentFirst = 1

			'Set the header and footer for the first page.
			sheet.PageSetup.FirstHeaderString = "Different First page"
			sheet.PageSetup.FirstFooterString = "Different First footer"

			'Set the other pages' header and footer. 
			sheet.PageSetup.LeftHeader = "Demo of Spire.XLS"
			sheet.PageSetup.CenterFooter = "Footer by Spire.XLS"

			Dim result As String = "Result-AddDifferentHeaderFooterForTheFirstPage.xlsx"

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

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
