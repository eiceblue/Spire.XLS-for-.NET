Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace AutoFitBasedOnCellValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new workbook object.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            'Get the cell range for cell B8 in the worksheet.
            Dim cell As CellRange = worksheet.Range("B8")
            'Set the text value of the cell to "Welcome to Spire.XLS!".
            cell.Text = "Welcome to Spire.XLS!"

            'Get the style of the cell.
            Dim style As CellStyle = cell.Style
            'Set the font size to 16.
            style.Font.Size = 16
            'Set the font to bold.
            style.Font.IsBold = True

            'Adjust the column width to fit the content of the cell.
            cell.AutoFitColumns()
            'Adjust the row height to fit the content of the cell.
            cell.AutoFitRows()

            'Specify the file name for the output file.
            Dim outputFile As String = "Output.xlsx"

            'Save the workbook to the specified output file using Excel 2013 version.
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
            Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
