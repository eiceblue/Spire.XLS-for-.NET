Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetDataValidationOnSeparateSheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "SetDataValidationOnSeparateSheet.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SetDataValidationOnSeparateSheet.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            ' Set the text for cell B10 on sheet1
            sheet1.Range("B10").Text = "Here is a dataValidation example."

            ' Get the second worksheet (index 1) from the workbook
            Dim sheet2 As Worksheet = workbook.Worksheets(1)

            ' Enable the use of 3D ranges in data validation for the parent workbook
            sheet2.ParentWorkbook.Allow3DRangesInDataValidation = True

            ' Specify the data range for data validation on cell B11 of sheet1 as cells A1 to A7 on sheet2
            sheet1.Range("B11").DataValidation.DataRange = sheet2.Range("A1:A7")

            ' Save the modified workbook to a new Excel file with the filename "result.xlsx" and Excel 2013 format
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("result.xlsx")
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
