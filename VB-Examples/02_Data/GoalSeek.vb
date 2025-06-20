Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace GoalSeek
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first sheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the value of cell A1 to "100"
            sheet.Range("A1").Value = "100"

            ' Define the target cell as A2
            Dim targetCell As CellRange = sheet.Range("A2")

            ' Set the formula of the target cell to calculate the sum of A1 and B1
            targetCell.Formula = "=SUM(A1+B1)"

            ' Define the variable cell as B1
            Dim guessCell As CellRange = sheet.Range("B1")

            ' Create a new GoalSeek object
            Dim goalSeek As New Spire.Xls.GoalSeek()

            ' Try to calculate a trial solution for the goal seek, setting the target cell to 500 and using the guess cell
            Dim result As GoalSeekResult = goalSeek.TryCalculate(targetCell, 500, guessCell)

            ' Determine the solution for the goal seek
            result.Determine()

            ' Save the workbook to a file named "GoalSeek.xlsx" in Excel 2013 format
            workbook.SaveToFile("GoalSeek.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'view the document
            ExcelDocViewer("GoalSeek.xlsx")
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
