Imports Spire.Xls
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms

Namespace AllShapesToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a workbook
			Dim workbook As New Workbook()

			'Load an excel file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Shape.xlsx")

			'Get the first worksheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Save all shape to images
			Dim shapelist As New SaveShapeTypeOption()
			shapelist.SaveAll = True
			Dim images As List(Of Bitmap) = worksheet.SaveShapesToImage(shapelist)
			Dim index As Integer = 0

			' Save all images
			For Each img As Image In images
				Dim imageFileName As String = "Image_" & index & ".png"
				img.Save(imageFileName, ImageFormat.Png)
				index += 1
			Next img
			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            List<SkiaSharp.SKBitmap> images = worksheet.SaveShapesToImage(shapelist);
'            int index = 0;
'            foreach (SkiaSharp.SKBitmap img in images)
'            {      
'                SkiaSharp.SKImage image = SkiaSharp.SKImage.FromBitmap(img);
'                string filename = "Image_" + index + ".png";
'                FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);
'                image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100).SaveTo(fileStream);
'                index++;
'            }
'            

			' Dispose of the workbook object to release resources
			workbook.Dispose()
		End Sub

		Private Sub OutputViewer(ByVal filename As String)
			Try
				System.Diagnostics.Process.Start(filename)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

		End Sub
	End Class
End Namespace
