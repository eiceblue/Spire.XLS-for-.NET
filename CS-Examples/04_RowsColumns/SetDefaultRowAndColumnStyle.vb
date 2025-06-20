Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetDefaultRowAndColumnStyle
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new Workbook object.
            Dim workbook As New Workbook()

            'Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Create a cell style with "Mystyle" name.
            Dim style As CellStyle = workbook.Styles.Add("Mystyle")
            ' Set its color to Yellow.
            style.Color = Color.Yellow

            'Apply the above style to the default style of the first row.
            sheet.SetDefaultRowStyle(1, style)
            'Apply the above style to the default style of the first column.
            sheet.SetDefaultColumnStyle(1, style)


            'Specify the name for the resulting file.
            Dim result As String = "result.xlsx"

            'Save the workbook to a file with the specified name and Excel version (in this case, Excel 2010).
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
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
