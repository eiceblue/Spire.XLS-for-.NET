Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace VerifyProtectedWorksheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ProtectedWorksheet.xlsx")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Check if the worksheet is password protected
            Dim detect As Boolean = worksheet.IsPasswordProtected

            ' Create a StringBuilder to store the content
            Dim content As New StringBuilder()

            ' Format the result string indicating whether the first worksheet is password protected or not
            Dim result As String = String.Format("The first worksheet is password protected or not: " & detect)

            ' Append the result to the content StringBuilder
            content.AppendLine(result)

            ' Specify the output file path
            Dim outputFile As String = "Output.txt"

            ' Write the content to the output file
            File.WriteAllText(outputFile, content.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the output file
            Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
