Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace SetInternalMarginOfTextbox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a text box to the worksheet at the specified location and size
            Dim textbox As XlsTextBoxShape = TryCast(sheet.TextBoxes.AddTextBox(4, 2, 100, 300), XlsTextBoxShape)

            ' Set the text content of the text box
            textbox.Text = "Insert TextBox in Excel and set the margin for the text"

            ' Set the horizontal alignment of the text box
            textbox.HAlignment = CommentHAlignType.Center

            ' Set the vertical alignment of the text box
            textbox.VAlignment = CommentVAlignType.Center

            ' Set the left inner margin of the text box
            textbox.InnerLeftMargin = 1

            ' Set the right inner margin of the text box
            textbox.InnerRightMargin = 3

            ' Set the top inner margin of the text box
            textbox.InnerTopMargin = 1

            ' Set the bottom inner margin of the text box
            textbox.InnerBottomMargin = 1

            ' Specify the result file name for saving the modified workbook
            Dim result As String = "Result-SetInternalMarginOfTextbox.xlsx"

            ' Save the workbook to a file with the specified name and format (Excel 2013)
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
