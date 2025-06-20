Imports Spire.Xls

Namespace ShowOrHideGridLine

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample2.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            ' Get the second worksheet from the workbook
            Dim sheet2 As Worksheet = workbook.Worksheets(1)

            ' Hide gridlines on the first worksheet
            sheet1.GridLinesVisible = False

            ' Show gridlines on the second worksheet
            sheet2.GridLinesVisible = True

            ' Specify the output filename for the modified workbook
            Dim output As String = "ShowOrHideGridLine.xlsx"

            ' Save the workbook to a file in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
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
