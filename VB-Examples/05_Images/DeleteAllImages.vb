Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace DeleteAllImages
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_5.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Delete all images from the worksheet.  
            'Iterate through the pictures in reverse order.
            For i As Integer = sheet.Pictures.Count - 1 To 0 Step -1
                'Remove each picture from the worksheet.
                sheet.Pictures(i).Remove()
            Next i

            'Specify the name of the resulting file after deleting the images.
            Dim result As String = "Result-DeleteAllImages.xlsx"

            'Save the modified workbook to the specified output file using Excel 2013 format.
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
