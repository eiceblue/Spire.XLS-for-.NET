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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("../../../../../../Data/Template_Xls_1.xlsx")

            ' Retrieve the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the PageSetup object for the worksheet
            Dim pageSetup As PageSetup = sheet.PageSetup

            ' Set the property to print gridlines on the worksheet
            pageSetup.IsPrintGridlines = True

            ' Set the property to print row and column headings on the worksheet
            pageSetup.IsPrintHeadings = True

            ' Set the property to print in black and white
            pageSetup.BlackAndWhite = True

            ' Set the property to print comments in place
            pageSetup.PrintComments = PrintCommentType.InPlace

            ' Set the property to print in draft mode
            pageSetup.Draft = True

            ' Set the property to handle errors as "Not Available" during printing
            pageSetup.PrintErrors = PrintErrorsType.NA

            ' Specify the file name for the resulting Excel file
            Dim result As String = "Result-SetOtherPrintOptionsOfXlsFile.xlsx"

            ' Save the modified workbook to the specified file in Excel 2013 format
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
	End Class
End Namespace
