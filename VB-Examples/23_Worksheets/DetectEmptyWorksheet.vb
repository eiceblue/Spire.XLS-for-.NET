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
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

			'Get the first worksheet
			Dim worksheet1 As Worksheet = workbook.Worksheets(0)

			'Detect the first worksheet is empty or not
			Dim detect1 As Boolean = worksheet1.IsEmpty

			'Get the second worksheet
			Dim worksheet2 As Worksheet = workbook.Worksheets(1)

			'Detect the second worksheet is empty or not
			Dim detect2 As Boolean = worksheet2.IsEmpty

			'Create StringBuilder to save 
			Dim content As New StringBuilder()

			'Set string format for displaying
			Dim result As String = String.Format("The first worksheet is empty or not: " & detect1 & vbCrLf & "The second worksheet is empty or not: " & detect2)

			'Add result string to StringBuilder
			content.AppendLine(result)

			'String for output file 
			Dim outputFile As String = "Output.txt"

			'Save them to a txt file
			File.WriteAllText(outputFile, content.ToString())

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
