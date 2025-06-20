Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CopyOnlyFormulaValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            '  Create a new workbook object.
            Dim workbook As New Workbook()

            ' Load an Excel document named "CopyOnlyFormulaValue.xlsx" from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CopyOnlyFormulaValue.xlsx")

            ' Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Specify the copy option as OnlyCopyFormulaValue, which means only the formula values will be copied.
            Dim copyOptions As CopyRangeOptions = CopyRangeOptions.OnlyCopyFormulaValue

            ' Copy the range from A2 to C2 and paste it to A5 to C5 in the same worksheet, using the specified copy options.
            sheet.Copy(sheet.Range("A2:C2"), sheet.Range("A5:C5"), copyOptions)

            ' Save the modified workbook to a file named "result.xlsx" using Excel 2010 format.
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer("result.xlsx")
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
