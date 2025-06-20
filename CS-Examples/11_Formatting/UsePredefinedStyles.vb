Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace UsePredefinedStyles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a new cell style named "newStyle"
            Dim style As CellStyle = workbook.Styles.Add("newStyle")
            style.Font.FontName = "Calibri"
            style.Font.IsBold = True
            style.Font.Size = 15
            style.Font.Color = Color.CornflowerBlue

            ' Get the cell range B5
            Dim range As CellRange = sheet.Range("B5")

            ' Set the text of the cell and apply the "newStyle" to it
            range.Text = "Welcome to use Spire.XLS"
            range.CellStyleName = style.Name

            ' Auto-fit the columns in the range
            range.AutoFitColumns()

            ' Save the modified workbook to a new file
            Dim result As String = "UsePredefinedStyles_result.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
