Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace CroppedPositionOfPicture
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
			Dim sheet1 As Worksheet = workbook.Worksheets(0)

			'Get the image from the first sheet
			Dim picture As ExcelPicture = sheet1.Pictures(0)

			'Get the cropped position
			Dim left As Integer = picture.Left
			Dim top As Integer = picture.Top
			Dim width As Integer = picture.Width
			Dim height As Integer = picture.Height

			'Create StringBuilder to save 
			Dim content As New StringBuilder()

			'Set string format for displaying
			Dim displayString As String = String.Format("Crop position: Left " & left & vbCrLf & "Crop position: Top " & top & vbCrLf & "Crop position: Width " & width & vbCrLf & "Crop position: Height " & height)

			'Add result string to StringBuilder
			content.AppendLine(displayString)

			'String for .txt file 
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
