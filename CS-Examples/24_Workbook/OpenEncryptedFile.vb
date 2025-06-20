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
            ' Set the file path of the encrypted Excel file
            Dim filePath As String = "..\..\..\..\..\..\Data\EncryptedFile.xlsx"

            ' Create a StringBuilder to store the result messages
            Dim builder As New StringBuilder()

            ' Define an array of passwords
            Dim passwords() As String = {"password1", "password2", "password3", "1234"}

            ' Iterate through each password in the array
            For i As Integer = 0 To passwords.Length - 1

                Try
                    ' Create a new Workbook object
                    Dim workbook As New Workbook()

                    ' Set the password for opening the workbook
                    workbook.OpenPassword = passwords(i)

                    ' Load the encrypted Excel file into the workbook
                    workbook.LoadFromFile(filePath)

                    ' Release the resources used by the workbook
                    workbook.Dispose()

                    ' Append a success message to the builder
                    builder.AppendLine("Password = " & passwords(i) & " is correct." & " The encrypted Excel file opened successfully!")
                Catch ex As Exception
                    ' Append an error message to the builder
                    builder.AppendLine("Password = " & passwords(i) & " is not correct")
                End Try

            Next i

            ' Specify the output file path
            Dim result As String = "OpenEncryptedFile_out.txt"

            ' Write the contents of the builder to the output file
            File.WriteAllText(result, builder.ToString())

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
