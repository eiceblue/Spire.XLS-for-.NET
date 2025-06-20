Imports Spire.Xls

Namespace HideOrShowWorksheet

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample3.xlsx")

            ' Hide the worksheet with the name "Sheet1"
            workbook.Worksheets("Sheet1").Visibility = WorksheetVisibility.Hidden

            ' Make the first worksheet visible
            workbook.Worksheets(1).Visibility = WorksheetVisibility.Visible

            ' Specify the output file path
            Dim output As String = "HideOrShowWorksheet.xlsx"

            ' Save the modified workbook to a file in Excel 2013 format
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
