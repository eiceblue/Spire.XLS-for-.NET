Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace GetStyleSetStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the template file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Specify the range as cell B4 in the worksheet
            Dim range As CellRange = sheet.Range("B4")

            ' Get the style of the specified range
            Dim style As CellStyle = range.Style

            ' Set the font properties of the style
            style.Font.FontName = "Calibri"
            style.Font.IsBold = True
            style.Font.Size = 15
            style.Font.Color = Color.CornflowerBlue

            ' Apply the modified style to the range
            range.Style = style

            ' Save the modified workbook to a new file named "UseGetStyleSetStyle_result.xlsx" in Excel 2010 format
            Dim result As String = "UseGetStyleSetStyle_result.xlsx"
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
