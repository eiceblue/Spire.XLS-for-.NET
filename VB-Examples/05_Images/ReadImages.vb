Imports System
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data

Imports Spire.Xls

Namespace ReadImages

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			'Create a Workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the first image
			Dim pic As ExcelPicture = sheet.Pictures(0)

			' Show Picture in the PictureBox
			Using frm1 As New Form()
				Dim pic1 As New PictureBox()
				pic1.Image = pic.Picture
				frm1.Width = pic.Picture.Width
				frm1.Height = pic.Picture.Height
				frm1.StartPosition = FormStartPosition.CenterParent
				pic1.Dock = DockStyle.Fill
				frm1.Controls.Add(pic1)
				frm1.ShowDialog()
			End Using

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            SkiaSharp.SKImage image = SkiaSharp.SKImage.FromBitmap(pic.Picture);
'            FileStream fileStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write);
'            image.Encode(SkiaSharp.SKEncodedImageFormat.Jpeg, 100).SaveTo(fileStream);
'            fileStream.Flush();
'            

			' Dispose of the workbook object to release resources
			workbook.Dispose()
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub


	End Class
End Namespace
