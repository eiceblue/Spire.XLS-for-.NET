Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace WriteImages

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Load an Excel document from the specified file.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WriteImages.xlsx")

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Add an image to the specified cell at position (14, 5) using the image file "SpireXls.png".
            sheet.Pictures.Add(14, 5, "..\..\..\..\..\..\Data\SpireXls.png")

            'Save the workbook to the file "Output.xlsx" using Excel 2010 format.
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
