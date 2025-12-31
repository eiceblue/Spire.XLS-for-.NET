Imports Spire.Xls
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms

Namespace GroupShapeToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			' Create a workbook
			Dim workbook As New Workbook()

			' Load an excel file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\GroupShapeToImage.xlsx")

			' Get the first worksheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Save to image
			Dim saveShapeTypeOption As New SaveShapeTypeOption()
			saveShapeTypeOption.SaveGroupShape = True
			Dim images As List(Of Bitmap) = worksheet.SaveShapesToImage(saveShapeTypeOption)
			For i As Integer = 0 To images.Count - 1
				Dim imageFile As String = String.Format("Image-{0}.png", i)
				images(i).Save(imageFile, ImageFormat.Png)
			Next i

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            List<SkiaSharp.SKBitmap> images = worksheet.SaveShapesToImage(saveShapeTypeOption);
'            for (int i = 0; i < images.Count; i++)
'            { 
'                SkiaSharp.SKImage image = SkiaSharp.SKImage.FromBitmap(images[i]);
'                String imageFile = string.Format("Image-{0}.png", i);
'                FileStream fileStream = new FileStream(imageFile, FileMode.Create, FileAccess.Write);
'                image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100).SaveTo(fileStream);
'            }
'            

			workbook.Dispose()
		End Sub
		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				System.Diagnostics.Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
