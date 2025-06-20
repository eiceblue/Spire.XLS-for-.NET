Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace CopyRows
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object to create a workbook.
            Dim workbook As New Workbook()

            'Load the Excel file, which includes a pivot table, from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Copying.xls")
            'Get the second worksheet from the workbook.
            Dim sheet1 As Worksheet = workbook.Worksheets(1)
            'Get the first worksheet from the workbook.
            Dim sheet2 As Worksheet = workbook.Worksheets(0)

            'Copy the content of the first row to the third row in the same sheet.
            sheet1.Copy(sheet1.Rows(0), sheet1.Rows(2), True, True, True)


            'Copy the content of the first row in sheet1 to the second row in sheet2.
            sheet1.Copy(sheet1.Rows(0), sheet2.Rows(1), True, True, True)

            'Specify the name for the resulting file.
            Dim result As String = "CopyRows_result.xlsx"

            'Save the workbook to a file with the specified name and Excel version (in this case, Excel 2010).
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
