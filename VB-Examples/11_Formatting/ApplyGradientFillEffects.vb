Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core

Namespace ApplyGradientFillEffects
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Set the Excel version to 2010
            workbook.Version = ExcelVersion.Version2010

            ' Get the first Worksheet from the Workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the CellRange object for cell B5
            Dim range As CellRange = sheet.Range("B5")

            ' Set the row height of the range to 50
            range.RowHeight = 50

            ' Set the column width of the range to 30
            range.ColumnWidth = 30

            ' Set the text in the range to "Hello"
            range.Text = "Hello"

            ' Set the horizontal alignment of the range to center
            range.Style.HorizontalAlignment = HorizontalAlignType.Center

            ' Set the fill pattern of the range to Gradient
            range.Style.Interior.FillPattern = ExcelPatternType.Gradient

            ' Set the fore color of the gradient to RGB(255, 255, 255)
            range.Style.Interior.Gradient.ForeColor = Color.FromArgb(255, 255, 255)

            ' Set the back color of the gradient to RGB(79, 129, 189)
            range.Style.Interior.Gradient.BackColor = Color.FromArgb(79, 129, 189)

            ' Apply a two-color horizontal gradient shading effect to the gradient
            range.Style.Interior.Gradient.TwoColorGradient(GradientStyleType.Horizontal, GradientVariantsType.ShadingVariants1)

            ' Specify the filename for the saved workbook
            Dim result As String = "ApplyGradientFillEffects_result.xlsx"

            ' Save the workbook to the specified file with Excel 2010 format
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
