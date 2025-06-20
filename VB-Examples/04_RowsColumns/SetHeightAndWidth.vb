Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetHeightAndWidth
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new Workbook object.
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the Workbook object.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SetHeightAndWidth.xls")

            ' Get the first worksheet from the Workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Set the fourth column width to 30.
            worksheet.SetColumnWidth(4, 30)
            ' Set the fourth row height to 30.
            worksheet.SetRowHeight(4, 30)
            'Specify the name for the result file.
            Dim result As String = "SetHeightAndWidth_out.xlsx"
            'Save the workbook to a file with the specified name and Excel version (in this case, Excel 2010).
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
