Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace SetOtherPrintingOptions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("../../../../../../Data/Template_Xls_1.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the reference of the PageSetup of the worksheet.
			Dim pageSetup As PageSetup = sheet.PageSetup

			'Allow to print gridlines.
			pageSetup.IsPrintGridlines = True

			'Allow to print row/column headings.
			pageSetup.IsPrintHeadings = True

			'Allow to print worksheet in black & white mode.
			pageSetup.BlackAndWhite = True

			'Allow to print comments as displayed on worksheet.
			pageSetup.PrintComments = PrintCommentType.InPlace

			'Allow to print worksheet with draft quality.
			pageSetup.Draft = True

			'Allow to print cell errors as N/A.
			pageSetup.PrintErrors = PrintErrorsType.NA

			Dim result As String = "Result-SetOtherPrintOptionsOfXlsFile.xlsx"

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
