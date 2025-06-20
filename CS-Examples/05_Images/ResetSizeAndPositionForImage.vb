Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ResetSizeAndPositionForImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Add a picture to the worksheet at position (1, 1) with the specified image file path.
            Dim picture As ExcelPicture = sheet.Pictures.Add(1, 1, "SpireXls.png")

            'Set the width of the picture to 200 units.
            picture.Width = 200

            'Set the height of the picture to 200 units.
            picture.Height = 200

            'Set the distance of the left edge of the picture from the left side of the worksheet to 200 units.
            picture.Left = 200

            'Set the distance of the top edge of the picture from the top of the worksheet to 100 units.
            picture.Top = 100

            'Specify the filename to save the modified workbook.
            Dim result As String = "Result-ResetSizeAndPositionForImage.xlsx"

            'Save the workbook to the specified file using Excel 2013 format.
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
            ExcelDocViewer(result)
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
