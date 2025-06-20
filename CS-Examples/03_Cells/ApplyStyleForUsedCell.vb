Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ApplyStyleForUsedCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()
            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")

            'Creates a new cell style with the name "Mystyle".
            Dim cellStyle As CellStyle = workbook.Styles.Add("Mystyle")
            'Sets the cell style color to transparent.
            cellStyle.Color = Color.Transparent
            'Sets the border color of the cell style to black.
            cellStyle.Borders.KnownColor = ExcelColors.Black
            'Sets the border line style of the cell style to thin.
            cellStyle.Borders.LineStyle = LineStyleType.Thin
            'Sets the diagonal down border line style of the cell style to none.
            cellStyle.Borders(BordersLineType.DiagonalDown).LineStyle = LineStyleType.None
            'Sets the diagonal up border line style of the cell style to none.
            cellStyle.Borders(BordersLineType.DiagonalUp).LineStyle = LineStyleType.None

            'Iterates through each worksheet in the workbook.
            For Each worksheet As Worksheet In workbook.Worksheets
                'Applies the cell style to the used cells in the worksheet, excluding empty cells.
                worksheet.ApplyStyle(cellStyle, False, False)
            Next worksheet
            'Specifies the filename for the resulting Excel file.
            Dim result As String = "ApplyStyle_result.xlsx"

            'Saves the modified workbook to a file with the specified filename and Excel version.
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
