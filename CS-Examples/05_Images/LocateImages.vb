Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace LocateImages
	Partial Public Class Form1
		Inherits Form
		Public Sub New()

			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new workbook object.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\LocateImages.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Retrieve the first picture from the worksheet.
            Dim pic As ExcelPicture = sheet.Pictures(0)

            'Set the left column offset of the picture to 300, adjusting its horizontal position.
            pic.LeftColumnOffset = 300

            'Set the top row offset of the picture to 300, adjusting its vertical position.
            pic.TopRowOffset = 300

            'Save the modified workbook to the specified output file using Excel 2010 format.
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("Output.xlsx")
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
