Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace DaysFormula
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from a file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_12.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the formula of cell C4 to calculate the number of days between dates in cells A8 and A1
            sheet.Range("C4").Formula = "=DAYS(A8,A1)"

            ' Calculate all the formulas in the workbook
            workbook.CalculateAllValue()

            ' Save the modified workbook to a new file
            Dim result As String = "DaysFormula_result.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2016)
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
