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
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")

			'String for output file 
			Dim outputFile As String = "Output.tiff"

			'Convert workbook to Tiff
			JoinTiffImages(ToImage(workbook), outputFile, EncoderValue.CompressionLZW)

			'Launching the output file.
			Viewer(outputFile)
		End Sub

		Private Shared Function ToImage(ByVal workbook As Workbook) As Image()
			'Get the worksheet count of workbook
			Dim workSheetNo As Integer = workbook.Worksheets.Count

			'Create an array
			Dim images(workSheetNo - 1) As Image

			'Save worksheet to image and add the array
			For i As Integer = 0 To workSheetNo - 1
				Dim workSheet As Worksheet = workbook.Worksheets(i)
				Dim output As String = String.Format("result{0}.jpg",i+1)
				workSheet.SaveToImage(output)
				Dim image As Image= Image.FromFile(output)
				images(i) = image
			Next i
			Return images
		End Function

		Private Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
			Dim encoders() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders()
			For j As Integer = 0 To encoders.Length - 1
				If encoders(j).MimeType = mimeType Then
					Return encoders(j)
				End If
			Next j
			Throw New Exception(mimeType & " mime type not found in ImageCodecInfo")
		End Function

		Public Shared Sub JoinTiffImages(ByVal images() As Image, ByVal outFile As String, ByVal compressEncoder As EncoderValue)
			'Use the save encoder
			Dim enc As Encoder = Encoder.SaveFlag
			Dim ep As New EncoderParameters(2)
			ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.MultiFrame))
			ep.Param(1) = New EncoderParameter(Encoder.Compression, CLng(compressEncoder))
			Dim pages As Image = images(0)
			Dim frame As Integer = 0
			Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")
			For Each img As Image In images
				If frame = 0 Then
					pages = img
					'save the first frame
					pages.Save(outFile, info, ep)

				Else
					'save the intermediate frames
					ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.FrameDimensionPage))

					pages.SaveAdd(img, ep)
				End If
				If frame = images.Length - 1 Then
					'flush and close.
					ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.Flush))
					pages.SaveAdd(ep)
				End If
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
