Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace PictureRefRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PictureRefRange.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Set the value of cell A1 in the worksheet to "Spire.XLS".
            sheet.Range("A1").Value = "Spire.XLS"

            'Set the value of cell B3 in the worksheet to "E-iceblue".
            sheet.Range("B3").Value = "E-iceblue"

            'Access the first picture in the worksheet.
            Dim picture As ExcelPicture = sheet.Pictures(0)

            'Specify the range of cells (A1 to B3) that the picture should be anchored to.
            picture.RefRange = "A1:B3"

            'Specify the file name for the output Excel file.
            Dim result As String = "PictureRefRange_out.xlsx"

            'Save the workbook to the file "PictureRefRange_out.xlsx" in Excel 2013 format.
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()
            'Launch the Excel file
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
