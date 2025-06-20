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
            ' Specify the file name of the workbook to decrypt
            Dim fileName As String = "..\..\..\..\..\..\Data\DecryptWorkbook.xlsx"

            ' Check if the workbook is password-protected
            Dim value As Boolean = Workbook.IsPasswordProtected(fileName)

            ' If the workbook is password-protected, perform the following operations
            If value Then

                ' Create a new Workbook object
                Dim workbook As New Workbook()

                ' Set the open password for the workbook
                workbook.OpenPassword = "eiceblue"

                ' Load the password-protected workbook
                workbook.LoadFromFile(fileName)

                ' Remove the protection from the workbook
                workbook.UnProtect()

                ' Save the decrypted workbook to a new file with Excel Version 2010 format
                workbook.SaveToFile("DecryptWorkbook_result.xlsx", ExcelVersion.Version2010)

                ' Release the resources used by the workbook
                workbook.Dispose()
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
