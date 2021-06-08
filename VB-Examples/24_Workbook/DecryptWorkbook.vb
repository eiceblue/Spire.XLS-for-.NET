Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace DecryptWorkbook
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim fileName As String = "..\..\..\..\..\..\Data\DecryptWorkbook.xlsx"

			'Detect if the Excel workbook is password protected.
			Dim value As Boolean = Workbook.IsPasswordProtected(fileName)

			If value Then
				'Load a file with the password specified
				Dim workbook As New Workbook()
				workbook.OpenPassword = "eiceblue"
				workbook.LoadFromFile(fileName)

				'Decrypt workbook
				workbook.UnProtect()

				'Save the document
				workbook.SaveToFile("DecryptWorkbook_result.xlsx", ExcelVersion.Version2010)
			End If

			ExcelDocViewer("DecryptWorkbook_result.xlsx")
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
