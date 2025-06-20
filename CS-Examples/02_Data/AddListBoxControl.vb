Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddListBoxControl
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Excel workbook object.
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the value of cell A7 to "Beijing"
            sheet.Range("A7").Text = "Beijing"
            ' Set the value of cell A8 to "New York".
            sheet.Range("A8").Text = "New York"
            ' Set the value of cell A9 to "ChengDu".
            sheet.Range("A9").Text = "ChengDu"
            ' Set the value of cell A10 to "Paris".
            sheet.Range("A10").Text = "Paris"
            ' Set the value of cell A11 to "Boston".
            sheet.Range("A11").Text = "Boston"
            ' Set the value of cell A12 to "London".
            sheet.Range("A12").Text = "London"
            ' Set the value of cell C13 to "City :".
            sheet.Range("C13").Text = "City :"
            ' Apply bold font style to the text in cell C13.
            sheet.Range("C13").Style.Font.IsBold = True

            ' Add a listbox control at position (13, 4) with width 100 and height 80.
            Dim listBox As IListBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80)
            ' Set the selection type of the listbox to single.
            listBox.SelectionType = SelectionType.Single
            ' Set the selected index of the listbox to 2 (selects the third item).
            listBox.SelectedIndex = 2
            ' Enable 3D shading for the listbox.
            listBox.Display3DShading = True
            ' Set the data range for the listbox to cells A7 to A12.
            listBox.ListFillRange = sheet.Range("A7:A12")

            ' Specify the output file name.
            Dim output As String = "InsertListBoxControl_out.xlsx"
            ' Save the modified workbook to the specified file with Excel 2013 format.
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
