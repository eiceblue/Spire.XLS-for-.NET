Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace RemoveRowBasedOnKeyword
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()
            'Loads the Excel document from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorkbookToHTML.xlsx")
            'Retrieves the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Find the string "Address" in the worksheet.
            Dim cr As CellRange = sheet.FindString("Address", False, False)

            'Delete the row which includes the found string.
            sheet.DeleteRow(cr.Row)

            'Save the modified workbook to a file named "RemoveRowBasedOnKeyword.xlsx" in Excel 2010 format.
            workbook.SaveToFile("RemoveRowBasedOnKeyword.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer("RemoveRowBasedOnKeyword.xlsx")
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
