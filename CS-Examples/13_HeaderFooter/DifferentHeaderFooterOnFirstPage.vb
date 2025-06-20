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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text in cell A1 to "Hello World"
            sheet.Range("A1").Text = "Hello World"

            ' Set the text in cell F30 to "Hello World"
            sheet.Range("F30").Text = "Hello World"

            ' Set the text in cell G150 to "Hello World"
            sheet.Range("G150").Text = "Hello World"

            ' Enable different header and footer for the first page only
            sheet.PageSetup.DifferentFirst = 1

            ' Set the header text for the first page
            sheet.PageSetup.FirstHeaderString = "Different First page"

            ' Set the footer text for the first page
            sheet.PageSetup.FirstFooterString = "Different First footer"

            ' Set the left header text for all pages
            sheet.PageSetup.LeftHeader = "Demo of Spire.XLS"

            ' Set the center footer text for all pages
            sheet.PageSetup.CenterFooter = "Footer by Spire.XLS"

            ' Specify the file name for the resulting workbook
            Dim result As String = "Result-AddDifferentHeaderFooterForTheFirstPage.xlsx"

            ' Save the workbook to a file with the specified name, in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
