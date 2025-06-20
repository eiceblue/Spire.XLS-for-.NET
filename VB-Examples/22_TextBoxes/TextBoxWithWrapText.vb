Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes
Imports System.ComponentModel
Imports System.Text

Namespace TextBoxWithWrapText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class
            Dim workbook As New Workbook()

            ' Load an existing workbook from a file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\TextBoxSampleB.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Try to cast the first TextBox shape on the worksheet to XlsTextBoxShape
            Dim shape As XlsTextBoxShape = TryCast(sheet.TextBoxes(0), XlsTextBoxShape)

            ' Set the IsWrapText property of the TextBox shape to true
            shape.IsWrapText = True

            ' Specify the output file name
            Dim output As String = "TextBoxWithWrapText.xlsx"

            ' Save the modified workbook to a file in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer(output)
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
