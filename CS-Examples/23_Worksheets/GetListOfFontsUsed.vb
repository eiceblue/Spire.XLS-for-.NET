Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports System.IO

Namespace GetListOfFontsUsed
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz.xlsx")

            ' Create a new list to store Excel fonts
            Dim fonts As New List(Of ExcelFont)()

            ' Iterate through each worksheet in the workbook
            For Each sheet As Worksheet In workbook.Worksheets

                ' Iterate through each row in the current sheet
                For r As Integer = 0 To sheet.Rows.Length - 1

                    ' Iterate through each cell in the current row
                    For c As Integer = 0 To sheet.Rows(r).CellList.Count - 1

                        ' Add the font of the current cell to the list
                        fonts.Add(sheet.Rows(r).CellList(c).Style.Font)

                    Next c
                Next r
            Next sheet

            ' Create a StringBuilder object to store font information
            Dim strB As New StringBuilder()

            ' Iterate through each font in the list
            For Each font As ExcelFont In fonts

                ' Append the font name and size to the StringBuilder object
                strB.AppendLine(String.Format("FontName:{0}; FontSize:{1}", font.FontName, font.Size))

            Next font

            ' Specify the result file path
            Dim result As String = "GetListOfFontsUsed_result.txt"

            ' Write the font information to the result file
            File.WriteAllText(result, strB.ToString())

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
