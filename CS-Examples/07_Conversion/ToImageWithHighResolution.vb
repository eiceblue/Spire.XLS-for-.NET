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
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ConversionSample1.xlsx")

            ' Get the first worksheet from the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Use a MemoryStream to temporarily store the EMF (Enhanced Metafile) image data.
            Using ms As New MemoryStream()
                ' Convert the worksheet to an EMF stream starting from row 1, column 1 and ending at the last row and last column.
                worksheet.ToEMFStream(ms, 1, 1, worksheet.LastRow, worksheet.LastColumn)

                ' Create an Image object from the stream.
                Dim image As Image = image.FromStream(ms)

                ' Reset the resolution of the image to 300 DPI (dots per inch).
                Dim images As Bitmap = ResetResolution(TryCast(image, Metafile), 300)

                ' Specify the output file name for the JPEG image.
                Dim output As String = "ToImage.jpg"

                ' Save the image as a JPEG file.
                images.Save(output, ImageFormat.Jpeg)
            End Using
            ' Release the resources used by the workbook
            workbook.Dispose()

        End Sub

        ' Function: ResetResolution
        ' Description: Resizes a Metafile to a specified resolution and returns a Bitmap object.
        ' Parameters:
        '   - mf: The input Metafile that needs to be resized.
        '   - resolution: The desired resolution for the output Bitmap.
        ' Returns: A Bitmap object with the resized image.
        Private Function ResetResolution(ByVal mf As Metafile, ByVal resolution As Single) As Bitmap
            ' Calculate the width of the resized image based on the original width and resolution ratio.
            Dim width As Integer = CInt(Fix(mf.Width * resolution / mf.HorizontalResolution))

            ' Calculate the height of the resized image based on the original height and resolution ratio.
            Dim height As Integer = CInt(Fix(mf.Height * resolution / mf.VerticalResolution))

            ' Create a new Bitmap with the calculated width and height.
            Dim bmp As New Bitmap(width, height)

            ' Set the resolution of the Bitmap object to match the desired resolution.
            bmp.SetResolution(resolution, resolution)

            ' Create a Graphics object from the Bitmap.
            Dim g As Graphics = Graphics.FromImage(bmp)

            ' Draw the original Metafile onto the Graphics object at position (0, 0).
            g.DrawImage(mf, 0, 0)

            ' Dispose the Graphics object since it is no longer needed.
            g.Dispose()

            ' Return the resized Bitmap object.
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
