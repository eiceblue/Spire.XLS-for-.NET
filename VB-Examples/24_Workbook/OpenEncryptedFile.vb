Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace OpenEncryptedFile
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'File path
			Dim filePath As String = "..\..\..\..\..\..\Data\EncryptedFile.xlsx"

			'Create string builder
			Dim builder As New StringBuilder()

			Dim passwords() As String = { "password1", "password2", "password3", "1234" }
			For i As Integer = 0 To passwords.Length - 1
				Try
					'Create a workbook
					Dim workbook As New Workbook()

					'Open password
					workbook.OpenPassword = passwords(i)

					'Load the document
					workbook.LoadFromFile(filePath)

					builder.AppendLine("Password = " & passwords(i) & " is correct." & " The encrypted Excel file opened successfully!")
				Catch ex As Exception
					builder.AppendLine("Password = " & passwords(i) & "  is not correct")
				End Try
			Next i

			'Save to txt file
			Dim result As String = "OpenEncryptedFile_out.txt"
			File.WriteAllText(result,builder.ToString())

			'Launch the file
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
