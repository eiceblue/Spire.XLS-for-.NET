Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddScrollBarControl
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Excel workbook object.
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

            ' Get the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the value of cell B10 to 1.
            sheet.Range("B10").Value2 = 1
            ' Apply bold font style to the text in cell B10.
            sheet.Range("B10").Style.Font.IsBold = True

            ' Add a scroll bar control at position (10, 3) with width 150 and height 20.
            Dim scrollBar As IScrollBarShape = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20)
            ' Link the value of the scroll bar to cell B10.
            scrollBar.LinkedCell = sheet.Range("B10")
            ' Set the minimum value of the scroll bar to 1.
            scrollBar.Min = 1
            ' Set the maximum value of the scroll bar to 150.
            scrollBar.Max = 150
            ' Set the incremental change value when scrolling to 1.
            scrollBar.IncrementalChange = 1
            ' Enable 3D shading for the scroll bar.
            scrollBar.Display3DShading = True

            ' Specify the output file name.
            Dim output As String = "AddScrollBarControl_out.xlsx"
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
