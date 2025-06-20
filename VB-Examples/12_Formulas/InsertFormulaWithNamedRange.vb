Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core


Namespace InsertFormulaWithNamedRange
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the value of cell A1 in the sheet to "1"
            sheet.Range("A1").Value = "1"

            ' Set the value of cell A2 in the sheet to "1"
            sheet.Range("A2").Value = "1"

            ' Create a new named range object and add it to the workbook's named ranges collection
            Dim NamedRange As INamedRange = workbook.NameRanges.Add("NewNamedRange")

            ' Set the local name of the named range to "=SUM(A1+A2)"
            NamedRange.NameLocal = "=SUM(A1+A2)"

            ' Set the formula of cell C1 in the sheet to "NewNamedRange"
            sheet.Range("C1").Formula = "NewNamedRange"

            ' Specify the file name for saving the workbook as "result.xlsx"
            Dim result As String = "result.xlsx"

            ' Save the workbook to a file with the specified name and Excel version
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
