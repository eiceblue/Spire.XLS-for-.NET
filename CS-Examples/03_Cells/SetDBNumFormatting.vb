Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetDBNumFormatting
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()
            'Creates one empty worksheet in the workbook.
            workbook.CreateEmptySheets(1)

            'Retrieves the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Sets the value of cell A1 to 123.
            sheet.Range("A1").Value2 = 123
            'Sets the value of cell A2 to 456.
            sheet.Range("A2").Value2 = 456
            'Sets the value of cell A3 to 789.
            sheet.Range("A3").Value2 = 789

            'Specifies the range of cells from A1 to A3.
            Dim range As CellRange = sheet.Range("A1:A3")

            'Sets the number format of the range to a specified DBNum format.
            range.NumberFormat = "[DBNum2][$-804]General"

            'Adjusts the width of the columns in the range to fit the content.
            range.AutoFitColumns()

            'Specifies the name of the output file.
            Dim output As String = "SetDBNumFormatting_out.xlsx"
            'Saves the workbook to the specified file in Excel 2013 format.
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
            ExcelDocViewer(output)
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
