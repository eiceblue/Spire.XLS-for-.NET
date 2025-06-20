Imports Spire.Xls

Namespace RemovePageBreak

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load a workbook from a specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PageBreak.xlsx")

            ' Get the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Clear all vertical page breaks in the worksheet.
            sheet.VPageBreaks.Clear()

            ' Remove the horizontal page break at index 0 in the worksheet.
            sheet.HPageBreaks.RemoveAt(0)

            ' Set the worksheet view mode to Preview.
            sheet.ViewMode = ViewMode.Preview

            ' Specify the output file name for saving the modified workbook.
            Dim output As String = "RemovePageBreak.xlsx"

            ' Save the workbook to the specified file path using Excel 2013 format.
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
            ExcelDocViewer(output)
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
