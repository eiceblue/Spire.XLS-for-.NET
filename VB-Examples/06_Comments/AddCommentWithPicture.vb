Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports Spire.Xls

Namespace AddCommentWithPicture

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			' Create a workbook
			Dim workbook As New Workbook()

			' Load file from disk
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Set value for the range
			sheet.Range("C6").Text = "E-iceblue"

			' Add the comment
			Dim comment As ExcelComment = sheet.Range("C6").AddComment()

			' Load the image file
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\Logo.png")

			' Fill the comment with a customized background picture
			comment.Fill.CustomPicture(image, "logo.png")

			' Set the height and width of comment
			comment.Height = image.Height
			comment.Width = image.Width

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            //Load the image file
'            SkiaSharp.SKBitmap image = SkiaSharp.SKBitmap.Decode(@"..\..\..\..\..\..\Data\Logo.png");
'            comment.Fill.CustomPicture(@"..\..\..\..\..\..\Data\Logo.png");
'            //Set the height and width of comment
'            comment.Height = image.Height;
'            comment.Width = image.Width;
'            

			comment.Visible = True

			' Specify the resulting file name.
			Dim output As String = "AddCommentWithPicture.xlsx"

			' Save the modified workbook to a file using Excel 201 format.
			workbook.SaveToFile(output, ExcelVersion.Version2010)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			'Launch the Excel file
			ExcelDocViewer(output)
		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				System.Diagnostics.Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub

	End Class
End Namespace
