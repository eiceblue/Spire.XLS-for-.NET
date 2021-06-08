Imports System.Drawing.Imaging
Imports System.IO

Imports Spire.Xls

Namespace ToImageWithHighResolution

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ConversionSample1.xlsx")

			'Get the worksheet you want to convert
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Convert the worksheet to EMF stream
			Using ms As New MemoryStream()
				worksheet.ToEMFStream(ms, 1, 1, worksheet.LastRow, worksheet.LastColumn)

				'Create an image from the EMF stream
				Dim image As Image = Image.FromStream(ms)
				Dim images As Bitmap = ResetResolution(TryCast(image, Metafile), 300)

				'Save the image in JPG file format
				Dim output As String = "ToImage.jpg"
				images.Save(output, ImageFormat.Jpeg)

				'Launch the Excel file
				ExcelDocViewer(output)
			End Using

		End Sub

		'A custom function to reset the image resolution
		Private Shared Function ResetResolution(ByVal mf As Metafile, ByVal resolution As Single) As Bitmap
			Dim width As Integer = CInt(Fix(mf.Width * resolution / mf.HorizontalResolution))
			Dim height As Integer = CInt(Fix(mf.Height * resolution / mf.VerticalResolution))
			Dim bmp As New Bitmap(width, height)
			bmp.SetResolution(resolution, resolution)
			Dim g As Graphics = Graphics.FromImage(bmp)
			g.DrawImage(mf, 0, 0)
			g.Dispose()
			Return bmp
		End Function

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
