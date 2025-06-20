Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace RemoveFormulasButKeepValues
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "RemoveFormulasButKeepValues.xlsx" from a specific path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\RemoveFormulasButKeepValues.xlsx")

            ' Iterate through each Worksheet in the Workbook
            For Each sheet As Worksheet In workbook.Worksheets

                ' Iterate through each CellRange in the current Worksheet
                For Each cell As CellRange In sheet.Range

                    ' Check if the cell contains a formula
                    If cell.HasFormula Then

                        ' Store the value calculated by the formula in a variable
                        Dim value As Object = cell.FormulaValue

                        ' Clear the content of the cell, removing the formula
                        cell.Clear(ExcelClearOptions.ClearContent)

                        ' Set the value of the cell to the stored value, retaining only the calculated result
                        cell.Value2 = value
                    End If
                Next cell
            Next sheet

            ' Specify the name for the resulting file after removing formulas but keeping values
            Dim result As String = "Result-RemoveFormulasButKeepValues.xlsx"

            ' Save the modified Workbook to a file with Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
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
