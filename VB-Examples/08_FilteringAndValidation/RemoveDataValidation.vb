Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace RemoveDataValidation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "RemoveDataValidation.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\RemoveDataValidation.xlsx")

            ' Declare and initialize an array of Rectangle objects with a size of 1
            Dim rectangles(0) As Rectangle

            ' Create a new Rectangle object and assign it to the first element of the rectangles array
            rectangles(0) = New Rectangle(0, 0, 1, 2)

            ' Remove the data validation within the specified rectangles on the first worksheet of the workbook
            workbook.Worksheets(0).DVTable.Remove(rectangles)

            ' Specify the output filename for the resulting workbook after removing data validation
            Dim result As String = "Result-RemoveDataValidation.xlsx"

            ' Save the modified workbook to a new Excel file with Excel 2013 format
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
