Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace SubTotalFormula
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set values for cells A1 to C3
            sheet.Range("A1").NumberValue = 1
            sheet.Range("A2").NumberValue = 2
            sheet.Range("A3").NumberValue = 3
            sheet.Range("B1").NumberValue = 4
            sheet.Range("B2").NumberValue = 5
            sheet.Range("B3").NumberValue = 6
            sheet.Range("C1").NumberValue = 7
            sheet.Range("C2").NumberValue = 8
            sheet.Range("C3").NumberValue = 9

            ' Set a formula in cell A5 to calculate subtotal using function SUBTOTAL(1, A1:C3)
            sheet.Range("A5").Formula = "=SUBTOTAL(1,A1:C3)"

            ' Set a formula in cell B5 to calculate subtotal using function SUBTOTAL(2, A1:C3)
            sheet.Range("B5").Formula = "=SUBTOTAL(2,A1:C3)"

            ' Set a formula in cell C5 to calculate subtotal using function SUBTOTAL(5, A1:C3)
            sheet.Range("C5").Formula = "=SUBTOTAL(5,A1:C3)"

            ' Calculate all formulas in the workbook
            workbook.CalculateAllValue()

            ' Specify the file name for the saved workbook
            Dim result As String = "SubtotalFormula_result.xlsx"

            ' Save the workbook to a file with Excel 2010 format
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
