Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace LinkToOtherSheetCell
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

            ' Define a range of cells with cell A1
            Dim range As CellRange = sheet.Range("A1")

            ' Create a hyperlink object and associate it with the defined range
            Dim hyperlink As HyperLink = sheet.HyperLinks.Add(range)

            ' Set the type of the hyperlink to a workbook
            hyperlink.Type = HyperLinkType.Workbook

            ' Set the display text for the hyperlink
            hyperlink.TextToDisplay = "Link to Sheet2 cell C5"

            ' Set the address of the hyperlink to Sheet2 cell C5
            hyperlink.Address = "Sheet2!C5"

            ' Define the output file name as "result.xlsx"
            Dim result As String = "result.xlsx"

            ' Save the workbook to the specified file path using Excel 2010 format
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
