Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace AddWatermark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'initialize a new instance of workbook and load the test file
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AddWatermark.xlsx")

			'Insert image in a header to mimic a watermark
			Dim font As Font = New Font("Arial", 40)
			Dim watermark As String = "Confidential"

			For Each sheet As Worksheet In workbook.Worksheets
				'Call DrawText() to create an image
				Dim imgWtrmrk As Image = DrawText(watermark, font, Color.LightCoral, Color.White, sheet.PageSetup.PageHeight, sheet.PageSetup.PageWidth)

				'Set image as left header image
				sheet.PageSetup.LeftHeaderImage = imgWtrmrk
				sheet.PageSetup.LeftHeader = "&G"

				'The watermark will only appear in this mode, it will disappear if the mode is normal
				sheet.ViewMode = ViewMode.Layout
			Next sheet

			'Save and Launch
			workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
			Process.Start("Output.xlsx")
		End Sub

		Private Shared Function DrawText(ByVal text As String, ByVal font As Font, ByVal textColor As Color, ByVal backColor As Color, ByVal height As Double, ByVal width As Double) As Image
			'Create a bitmap image with specified width and height
			Dim img As Image = New Bitmap(CInt(Fix(width)), CInt(Fix(height)))
			Dim drawing As Graphics = Graphics.FromImage(img)

			'Get the size of text
			Dim textSize As SizeF = drawing.MeasureString(text, font)

			'Set rotation point
			drawing.TranslateTransform((CInt(Fix(width)) - textSize.Width) / 2, (CInt(Fix(height)) - textSize.Height) / 2)

			'Rotate text
			drawing.RotateTransform(-45)

			'Reset translate transform    
			drawing.TranslateTransform(-(CInt(Fix(width)) - textSize.Width) / 2, -(CInt(Fix(height)) - textSize.Height) / 2)

			'Paint the background
			drawing.Clear(backColor)

			'Create a brush for the text
			Dim textBrush As Brush = New SolidBrush(textColor)

			'Draw text on the image at center position
			drawing.DrawString(text, font, textBrush, (CInt(Fix(width)) - textSize.Width) / 2, (CInt(Fix(height)) - textSize.Height) / 2)
			drawing.Save()
			Return img
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
