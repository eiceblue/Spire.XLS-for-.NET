Imports Spire.Xls

Namespace ApplyStyleToWorksheet

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook class
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a new cell style named "newStyle"
            Dim style As CellStyle = workbook.Styles.Add("newStyle")

            ' Set the background color of the style to LightBlue
            style.Color = Color.LightBlue

            ' Set the font color of the style to White
            style.Font.Color = Color.White

            ' Set the font size of the style to 15 points
            style.Font.Size = 15

            ' Make the font bold in the style
            style.Font.IsBold = True

            ' Apply the created style to the worksheet
            sheet.ApplyStyle(style)

            ' Specify the output file name
            Dim output As String = "ApplyStyleToWorksheet.xlsx"

            ' Save the workbook to the specified file in Excel 2013 format
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
