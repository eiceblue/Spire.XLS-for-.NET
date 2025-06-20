Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace CopyWorksheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to hold the source workbook
            Dim sourceWorkbook As New Workbook()

            ' Load an existing Excel file into the source workbook from a specified path
            sourceWorkbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

            ' Get the reference to the first worksheet in the source workbook
            Dim srcWorksheet As Worksheet = sourceWorkbook.Worksheets(0)

            ' Create a new Workbook object to hold the target workbook
            Dim targetWorkbook As New Workbook()

            ' Load an existing Excel file into the target workbook from a specified path
            targetWorkbook.LoadFromFile("..\..\..\..\..\..\Data\sample.xlsx")

            ' Add a new worksheet named "added" to the target workbook and get its reference
            Dim targetWorksheet As Worksheet = targetWorkbook.Worksheets.Add("added")

            ' Copy the contents (including formatting and formulas) of the source worksheet to the target worksheet
            targetWorksheet.CopyFrom(srcWorksheet)

            ' Specify the output file name for the modified target workbook
            Dim outputFile As String = "Output.xlsx"

            ' Save the modified target workbook to a new file with the specified name and Excel version (2013)
            targetWorkbook.SaveToFile(outputFile, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            sourceWorkbook.Dispose()

            ' Release the resources used by the workbook
            targetWorkbook.Dispose()

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
