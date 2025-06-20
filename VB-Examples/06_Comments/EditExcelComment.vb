Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace EditExcelComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Load an Excel document from the specified file.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_8.xlsx")

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Access the first comment in the worksheet.
            Dim comment As ExcelComment = sheet.Comments(0)

            'Set the text of the comment to "This comment has been edited by Spire.XLS."
            comment.Text = "This comment has been edited by Spire.XLS."

            'Specify the filename to save the modified workbook.
            Dim result As String = "Result-EditExcelComment.xlsx"

            'Save the workbook to the specified file using Excel 2013 format.
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
