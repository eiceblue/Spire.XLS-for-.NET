Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core
Imports System.IO
Imports System.Text

Namespace ExtractTextImageFromShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_5.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Extract text from the first shape and save to a txt file.
			Dim shape1 As IPrstGeomShape = sheet.PrstGeomShapes(2)
			Dim s As String = shape1.Text
			Dim sb As New StringBuilder()
			sb.AppendLine("The text in the third shape is: " & s)
			Dim result1 As String = "Result-ExtractTextAndImageFromShape.txt"
			File.WriteAllText(result1, sb.ToString())

			'Extract image from the second shape and save to a local folder.
			Dim shape2 As IPrstGeomShape = sheet.PrstGeomShapes(1)
			Dim image As Image = shape2.Fill.Picture
			Dim result2 As String = "Result-ExtractTextAndImageFromShape.png"
			image.Save(result2, System.Drawing.Imaging.ImageFormat.Png)

			'Launch the .txt file.
			ExcelDocViewer(result1)

			'Launch the image.
			ExcelDocViewer(result2)

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
