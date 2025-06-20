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
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from a file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_5.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Retrieve the third preset geometry shape from the worksheet
            Dim shape1 As IPrstGeomShape = sheet.PrstGeomShapes(2)

            ' Get the text content from shape1
            Dim s As String = shape1.Text

            ' Create a StringBuilder and append the extracted text
            Dim stringBuilder As New StringBuilder()
            stringBuilder.AppendLine("The text in the third shape is: " & s)

            ' Specify the resulting file name for saving the text content
            Dim result1 As String = "Result-ExtractTextAndImageFromShape.txt"

            ' Write the text content to a file
            File.WriteAllText(result1, stringBuilder.ToString())

            ' Retrieve the second preset geometry shape from the worksheet
            Dim shape2 As IPrstGeomShape = sheet.PrstGeomShapes(1)

            ' Get the image from the fill of shape2
            Dim image As Image = shape2.Fill.Picture

            ' Specify the resulting file name for saving the image
            Dim result2 As String = "Result-ExtractTextAndImageFromShape.png"

            ' Save the image to a file in PNG format
            image.Save(result2, System.Drawing.Imaging.ImageFormat.Png)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
