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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Get the PageSetup object for the worksheet
            Dim pageSetup As PageSetup = worksheet.PageSetup

            ' Set the print area to be printed on the worksheet (cells A1 to E19)
            pageSetup.PrintArea = "A1:E19"

            ' Set the columns to repeat on each printed page (columns A to E)
            pageSetup.PrintTitleColumns = "$A:$E"

            ' Set the rows to repeat at the top of each printed page (rows 1 and 2)
            pageSetup.PrintTitleRows = "$1:$2"

            ' Enable printing of gridlines on the worksheet
            pageSetup.IsPrintGridlines = True

            ' Enable printing of headings on the worksheet
            pageSetup.IsPrintHeadings = True

            ' Set the print mode to black and white
            pageSetup.BlackAndWhite = True

            ' Set the type of comments to be printed in place
            pageSetup.PrintComments = PrintCommentType.InPlace

            ' Set the print quality to 150 dots per inch
            pageSetup.PrintQuality = 150

            ' Set the type of errors to be displayed when printing if any occur
            pageSetup.PrintErrors = PrintErrorsType.NA

            ' Set the order of printing (over then down)
            pageSetup.Order = OrderType.OverThenDown

            ' Print the workbook's contents using the default printer
            workbook.PrintDocument.Print()

            ' Release the resources used by the workbook
            workbook.Dispose()
        End Sub
		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
