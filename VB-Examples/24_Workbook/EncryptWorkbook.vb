Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace EncryptWorkbook
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook and load a file
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\EncryptWorkbook.xlsx")

			'Protect Workbook with the password you want
			workbook.Protect("eiceblue")

			'Save the document and launch it
			workbook.SaveToFile("EncryptWorkbook_result.xlsx", ExcelVersion.Version2010)
			ExcelDocViewer("EncryptWorkbook_result.xlsx")
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
