Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace DetectEmptyWorksheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

            ' Get the reference to the first worksheet in the workbook
            Dim worksheet1 As Worksheet = workbook.Worksheets(0)

            ' Check if the first worksheet is empty or not and store the result in a Boolean variable
            Dim detect1 As Boolean = worksheet1.IsEmpty

            ' Get the reference to the second worksheet in the workbook
            Dim worksheet2 As Worksheet = workbook.Worksheets(1)

            ' Check if the second worksheet is empty or not and store the result in a Boolean variable
            Dim detect2 As Boolean = worksheet2.IsEmpty

            ' Create a StringBuilder object to hold the content
            Dim content As New StringBuilder()

            ' Create a formatted string with the results of the worksheet emptiness checks
            Dim result As String = String.Format("The first worksheet is empty or not: " & detect1 & vbCrLf & "The second worksheet is empty or not: " & detect2)

            ' Append the result to the content StringBuilder
            content.AppendLine(result)

            ' Specify the output file name for the text file
            Dim outputFile As String = "Output.txt"

            ' Write the content of the StringBuilder to the output file
            File.WriteAllText(outputFile, content.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
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
