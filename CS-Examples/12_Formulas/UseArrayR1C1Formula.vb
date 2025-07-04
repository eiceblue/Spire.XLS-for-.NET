﻿Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace UseArrayR1C1Formula
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

            ' Set numeric values for cells A1 to C3
            sheet.Range("A1").NumberValue = 1
            sheet.Range("A2").NumberValue = 2
            sheet.Range("A3").NumberValue = 3
            sheet.Range("B1").NumberValue = 4
            sheet.Range("B2").NumberValue = 5
            sheet.Range("B3").NumberValue = 6
            sheet.Range("C1").NumberValue = 7
            sheet.Range("C2").NumberValue = 8
            sheet.Range("C3").NumberValue = 9

            ' Set the text "Sum:" in cell B4
            sheet.Range("B4").Text = "Sum:"

            ' Align the text in cell B4 to the right
            sheet.Range("B4").Style.HorizontalAlignment = HorizontalAlignType.Right

            ' Write an array formula in cell C4 with R1C1 notation to calculate the sum of values in range A1:C3
            sheet.Range("C4").FormulaArrayR1C1 = "=SUM(R[-3]C[-2]:R[-1]C)"

            ' Calculate all the formulas in the workbook
            workbook.CalculateAllValue()

            ' Specify the name for the resulting file after applying array formulas with R1C1 notation
            Dim result As String = "UseArrayR1C1Formulas_result.xlsx"

            ' Save the modified Workbook to a file with Excel 2010 format
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
