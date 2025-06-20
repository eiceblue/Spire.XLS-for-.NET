Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace AddWorksheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AddWorksheet.xlsx")

            ' Add a new worksheet with the name "AddedSheet" to the workbook
            Dim sheet As Worksheet = workbook.Worksheets.Add("AddedSheet")

            ' Set the text of cell C5 in the added worksheet
            sheet.Range("C5").Text = "This is a new sheet."

            ' Save the modified workbook to a new file named "Output.xlsx" with Excel 2010 format
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Output.xlsx")
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
