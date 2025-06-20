Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace FormatCellsWithStyle
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")

            ' Create a new CellStyle object named "newStyle"
            Dim style As CellStyle = workbook.Styles.Add("newStyle")

            ' Set the color of the style to DarkGray
            style.Color = Color.DarkGray

            ' Set the font color of the style to White
            style.Font.Color = Color.White

            ' Set the font name of the style to "Times New Roman"
            style.Font.FontName = "Times New Roman"

            ' Set the font size of the style to 12
            style.Font.Size = 12

            ' Make the font bold in the style
            style.Font.IsBold = True

            ' Set the rotation angle of the style to 45 degrees
            style.Rotation = 45

            ' Set the horizontal alignment of the style to Center
            style.HorizontalAlignment = HorizontalAlignType.Center

            ' Set the vertical alignment of the style to Center
            style.VerticalAlignment = VerticalAlignType.Center

            ' Apply the style to the range A1:J1 in the first worksheet of the workbook
            workbook.Worksheets(0).Range("A1:J1").CellStyleName = style.Name

            ' Specify the filename for the resulting saved file
            Dim result As String = "result.xlsx"

            ' Save the workbook to the specified file in Excel 2010 format
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
