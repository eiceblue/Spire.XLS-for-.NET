Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace PageSetupForPrinting
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")

			'Get the first worksheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Specifying the print area
			Dim pageSetup As PageSetup = worksheet.PageSetup
			pageSetup.PrintArea = "A1:E19"

			'Define column A & E as title columns
			pageSetup.PrintTitleColumns = "$A:$E"

			'Define row numbers 1 as title rows
			pageSetup.PrintTitleRows = "$1:$2"

			'Allow to print with gridlines
			pageSetup.IsPrintGridlines = True

			'Allow to print with row/column headings
			pageSetup.IsPrintHeadings = True

			'Allow to print worksheet in black & white mode
			pageSetup.BlackAndWhite = True

			'Allow to print comments as displayed on worksheet
			pageSetup.PrintComments = PrintCommentType.InPlace

			'Set printing quality
			pageSetup.PrintQuality = 150

			'Allow to print cell errors as N/A
			pageSetup.PrintErrors = PrintErrorsType.NA

			'Set the printing order 
			pageSetup.Order = OrderType.OverThenDown

			workbook.PrintDocument.Print()
		End Sub
		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
