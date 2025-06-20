Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace GetConditionalFormatColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load a workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_13.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Define a CellRange object representing range A1:C1 in the worksheet
            Dim cRange As CellRange = sheet.Range("A1:C1")

            ' Get the color of the condition format applied to the range
            Dim color = cRange.GetConditionFormatsStyle().Color

            ' Display a message box showing the color of the condition format
            MessageBox.Show("The color of the condition format is " & color.ToString())

		End Sub


		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
