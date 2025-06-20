Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ReplaceFont
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Creates a new Excel workbook.
            Dim workbook As New Workbook()

            ' Loads an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")

            ' Retrieves the first worksheet in the workbook (index starts at 0).
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Adds a new cell style to the workbook and assigns it the name "newStyle".
            Dim newStyle As CellStyle = workbook.Styles.Add("newStyle")
            ' Sets the font name of the new style to "Arial Black".
            newStyle.Font.FontName = "Arial Black"
            ' Sets the font size of the new style to 14 points.
            newStyle.Font.Size = 14

            ' Initializes a variable to store the old style that will be replaced.
            Dim oldStyle As CellStyle = Nothing
            ' Iterates through all the styles in the workbook.
            For i As Integer = 0 To workbook.Styles.Count - 1
                ' Checks if the current style's font name is "Aleo".
                If workbook.Styles(i).Font.FontName = "Aleo" Then
                    ' Assigns the style of the cell D9 in the sheet to the oldStyle variable.
                    oldStyle = sheet.Range("D9").Style
                End If
            Next i

            ' Replaces all occurrences of "North America" in the sheet with the style specified by oldStyle, replacing it with "America" using the newStyle.
            sheet.ReplaceAll("North America", oldStyle, "America", newStyle)

            ' Saves the modified workbook to a new file named "ReplaceFont_out.xlsx" in the Excel 2013 format.
            workbook.SaveToFile("ReplaceFont_out.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'view the document
            ExcelDocViewer("ReplaceFont_out.xlsx")
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
