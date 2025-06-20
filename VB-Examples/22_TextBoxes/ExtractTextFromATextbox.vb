Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Shapes
Imports System.IO
Imports System.Text

Namespace ExtractTextFromATextbox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_5.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first TextBox shape from the worksheet and cast it to XlsTextBoxShape
            Dim shape As XlsTextBoxShape = TryCast(sheet.TextBoxes(0), XlsTextBoxShape)

            ' Create a StringBuilder to store the extracted text
            Dim content As New StringBuilder()
            content.AppendLine("The text extracted from the TextBox is: ")
            content.AppendLine(shape.Text)

            ' Specify the filename for the resulting text file
            Dim result As String = "Result-ExtractTextFromATextbox.txt"

            ' Write the content of the StringBuilder to the specified text file
            File.WriteAllText(result, content.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file.
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
