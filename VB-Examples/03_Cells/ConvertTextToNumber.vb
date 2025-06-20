Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace ConvertTextToNumber
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new workbook object.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Sample.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            'Convert the cells in range D2 to D8 from text string format to number format.
            worksheet.Range("D2:D8").ConvertToNumber()

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
