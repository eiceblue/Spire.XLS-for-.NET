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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\EncryptWorkbook.xlsx")

            ' Protect the workbook with a password
            workbook.Protect("eiceblue")

            ' Save the protected workbook to a new file with Excel Version 2010 format
            workbook.SaveToFile("EncryptWorkbook_result.xlsx", ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
