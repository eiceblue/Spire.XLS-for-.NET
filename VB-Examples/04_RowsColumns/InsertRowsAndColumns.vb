Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace InsertRowsAndColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()
            'Loads the Excel document from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\InsertRowsAndColumns.xls")
            'Retrieves the first worksheet from the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            'Inserting a row into the worksheet.
            worksheet.InsertRow(2)
            'Inserts the second column into the worksheet.
            worksheet.InsertColumn(2)
            'Inserts two rows from the fifth row into the worksheet.
            worksheet.InsertRow(5, 2)
            'Inserting two columns from the fifth column into the worksheet.
            worksheet.InsertColumn(5, 2)
            'Specifies the name of the resulting Excel file.
            Dim result As String = "InsertRowsAndColumns_out.xlsx"
            'Saves the modified workbook to a file with the specified name and Excel version.
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
