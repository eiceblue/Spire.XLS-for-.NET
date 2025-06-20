Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace SplitDataIntoMultipleColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SplitExcelDataIntoMultipleCols.xlsx")

            'Gets the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Declares an array to store the split text.
            Dim splitText() As String = Nothing
            'Declares a variable to store the text from each cell.
            Dim text As String = Nothing
            'Iterates through each row in the worksheet, excluding the last row.
            For i As Integer = 1 To sheet.LastRow - 1
                'Gets the text from the current cell in the first column.
                text = sheet.Range(i + 1, 1).Text
                'Splits the text into an array using space as the delimiter.
                splitText = text.Split(" "c)
                'Iterates through each element in the splitText array.
                For j As Integer = 0 To splitText.Length - 1
                    'Sets the value of the corresponding cell in the next column to the split text.
                    sheet.Range(i + 1, 1 + j + 1).Text = splitText(j)
                Next j
            Next i
            'Specifies the filename for the resulting Excel file.
            Dim result As String = "Result-SplitExcelDataIntoMultipleColumns.xlsx"

            'Saves the modified workbook to a file with the specified filename and Excel version.
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
