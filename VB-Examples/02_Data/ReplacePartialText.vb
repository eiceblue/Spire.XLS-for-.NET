Imports Spire.Xls


Namespace ReplacePartialText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Creates a new Excel workbook.
            Dim workbook As New Workbook()

            ' Retrieves the first worksheet in the workbook (index starts at 0).
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Sets the text value of cell A1 to "Hello World".
            sheet.Range("A1").Text = "Hello World"
            ' Adjusts the width of column A to fit the content.
            sheet.Range("A1").AutoFitColumns()

            ' Replaces the partial text "World" with "Spire" in the first cell in the worksheet.
            sheet.CellList(0).TextPartReplace("World", "Spire")

            ' Saves the modified workbook to a file named "replaced.xlsx" in the Excel 2016 format.
            workbook.SaveToFile("replaced.xlsx", ExcelVersion.Version2016)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
            ExcelDocViewer("replaced.xlsx")
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
