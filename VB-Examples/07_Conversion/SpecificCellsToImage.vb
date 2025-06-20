Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Drawing.Imaging

Namespace SpecificCellsToImage

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ConversionSample1.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Convert the specified range (1, 1, 7, 5) of the worksheet to an image and save it as PNG format
            sheet.ToImage(1, 1, 7, 5).Save("image1.png", ImageFormat.Png)

            ' Convert the specified range (8, 1, 15, 5) of the worksheet to an image and save it as JPEG format
            sheet.ToImage(8, 1, 15, 5).Save("image2.jpg", ImageFormat.Jpeg)

            ' Convert the specified range (17, 1, 23, 5) of the worksheet to an image and save it as BMP format
            sheet.ToImage(17, 1, 23, 5).Save("image3.bmp", ImageFormat.Bmp)
            ' Release the resources used by the workbook
            workbook.Dispose()
        End Sub
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
