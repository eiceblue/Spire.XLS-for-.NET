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
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from a file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AddWatermark.xlsx")

            ' Define the font and watermark text
            Dim font As Font = New Font("Arial", 40)
            Dim watermark As String = "Confidential"

            ' Loop through each worksheet in the workbook
            For Each sheet As Worksheet In workbook.Worksheets

                ' Draw the watermark image
                Dim imgWtrmrk As Image = DrawText(watermark, font, Color.LightCoral, Color.White, sheet.PageSetup.PageHeight, sheet.PageSetup.PageWidth)

                ' Set the watermark image as the left header image
                sheet.PageSetup.LeftHeaderImage = imgWtrmrk

                ' Set the left header to display the page number
                sheet.PageSetup.LeftHeader = "&G"

                ' Set the view mode to layout
                sheet.ViewMode = ViewMode.Layout
            Next sheet

            ' Save the modified workbook to a new file
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()
            Process.Start("Output.xlsx")
		End Sub

        Private Function DrawText(ByVal text As String, ByVal font As Font, ByVal textColor As Color, ByVal backColor As Color, ByVal height As Double, ByVal width As Double) As Image

            ' Create a new bitmap image with specified width and height
            Dim img As Image = New Bitmap(CInt(Fix(width)), CInt(Fix(height)))

            ' Create a Graphics object from the image
            Dim drawing As Graphics = Graphics.FromImage(img)

            ' Measure the size of the text using the specified font
            Dim textSize As SizeF = drawing.MeasureString(text, font)

            ' Translate the drawing origin to center the text
            drawing.TranslateTransform((CInt(Fix(width)) - textSize.Width) / 2, (CInt(Fix(height)) - textSize.Height) / 2)

            ' Rotate the drawing surface by -45 degrees
            drawing.RotateTransform(-45)

            ' Translate the drawing origin back to its original position
            drawing.TranslateTransform(-(CInt(Fix(width)) - textSize.Width) / 2, -(CInt(Fix(height)) - textSize.Height) / 2)

            ' Clear the drawing surface with the specified background color
            drawing.Clear(backColor)

            ' Create a brush for the text color
            Dim textBrush As Brush = New SolidBrush(textColor)

            ' Draw the text on the image
            drawing.DrawString(text, font, textBrush, (CInt(Fix(width)) - textSize.Width) / 2, (CInt(Fix(height)) - textSize.Height) / 2)

            ' Save the drawing operations
            drawing.Save()

            ' Return the resulting image with the watermark
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
