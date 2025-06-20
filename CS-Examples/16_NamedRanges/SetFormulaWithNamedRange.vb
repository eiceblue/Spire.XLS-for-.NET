Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace SetFormulaWithNamedRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a named range object and assign it to "MyNamedRange"
            Dim NamedRange As INamedRange = workbook.NameRanges.Add("MyNamedRange")
            NamedRange.RefersToRange = sheet.Range("B10:B12")

            ' Set the formula in cell B13 to calculate the sum of the named range
            sheet.Range("B13").Formula = "=SUM(MyNamedRange)"

            ' Set values in cells B10, B11, and B12
            sheet.Range("B10").Value2 = 10
            sheet.Range("B11").Value2 = 20
            sheet.Range("B12").Value2 = 30

            ' Save the modified workbook to a new file
            Dim result As String = "SetFormulaWithNamedRange_out.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
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
