Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.Drawing.Imaging

Namespace ToTiff
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")

            ' Specify the output filename for the TIFF image
            Dim outputFile As String = "Output.tiff"

            ' Join multiple images into a single TIFF image with LZW compression
            JoinTiffImages(ToImage(workbook), outputFile, EncoderValue.CompressionLZW)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
            Viewer(outputFile)
		End Sub

        Private Function ToImage(ByVal workbook As Workbook) As Image()
            ' Get the number of worksheets in the workbook
            Dim workSheetNo As Integer = workbook.Worksheets.Count

            ' Create an array to store the images
            Dim images(workSheetNo - 1) As Image

            ' Iterate through each worksheet and convert it to an image
            For i As Integer = 0 To workSheetNo - 1
                Dim workSheet As Worksheet = workbook.Worksheets(i)

                ' Generate the output filename for the image
                Dim output As String = String.Format("result{0}.jpg", i + 1)

                ' Save the worksheet as an image file
                workSheet.SaveToImage(output)

                ' Load the saved image into an Image object
                Dim image As Image = image.FromFile(output)

                ' Store the image in the array
                images(i) = image
            Next i

            ' Return the array of images
            Return images
        End Function

        Private Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
            ' Get the available image encoders
            Dim encoders() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders()

            ' Find the encoder that matches the specified MIME type
            For j As Integer = 0 To encoders.Length - 1
                If encoders(j).MimeType = mimeType Then
                    Return encoders(j)
                End If
            Next j

            ' Throw an exception if the specified MIME type is not found
            Throw New Exception(mimeType & " mime type not found in ImageCodecInfo")
        End Function

        Public Sub JoinTiffImages(ByVal images() As Image, ByVal outFile As String, ByVal compressEncoder As EncoderValue)
            ' Create an encoder for the SaveFlag parameter
            Dim enc As Encoder = Encoder.SaveFlag

            ' Create an EncoderParameters object to specify compression and multi-frame settings
            Dim ep As New EncoderParameters(2)
            ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.MultiFrame))
            ep.Param(1) = New EncoderParameter(Encoder.Compression, CLng(compressEncoder))

            ' Get the first image in the array
            Dim pages As Image = images(0)

            ' Initialize the frame counter
            Dim frame As Integer = 0

            ' Get the ImageCodecInfo for TIFF format
            Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")

            ' Process each image in the array
            For Each img As Image In images
                If frame = 0 Then
                    ' For the first image, save it as the main pages
                    pages = img
                    pages.Save(outFile, info, ep)
                Else
                    ' For subsequent images, add them as additional frames
                    ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.FrameDimensionPage))
                    pages.SaveAdd(img, ep)
                End If

                If frame = images.Length - 1 Then
                    ' For the last image, flush the encoder parameters to finalize the TIFF file
                    ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.Flush))
                    pages.SaveAdd(ep)
                End If

                ' Increment the frame counter
                frame += 1
            Next img
        End Sub
        Private Sub Viewer(ByVal fileName As String)
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
