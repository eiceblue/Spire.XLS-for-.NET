Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core
Imports System.Text
Imports System.IO

Namespace GetShapeLinkedCellRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            Using workbook As New Workbook()
                Dim stringBuilder As New StringBuilder()

                ' Load the Excel file
                workbook.LoadFromFile("..\..\..\..\..\..\Data\CellLinkedRangeLocal.xlsx")

                ' Get the first worksheet
                Dim sheet As Worksheet = workbook.Worksheets(0)

                ' Get the collection of preset geometric shapes
                Dim prstGeomShapeCollection As PrstGeomShapeCollection = sheet.PrstGeomShapes

                ' Get a specific shape named "Yesterday"
                Dim shape As IPrstGeomShape = prstGeomShapeCollection("Yesterday")

                ' Get the range address of the linked cell for the shape
                Dim cellAddress As String = shape.LinkedCell.RangeAddress

                ' Append the cell address to the string builder
                stringBuilder.Append(cellAddress & vbLf)

                ' Get another shape named "NewShapes"
                shape = prstGeomShapeCollection("NewShapes")

                ' Get the range address of the linked cell for the shape
                cellAddress = shape.LinkedCell.RangeAddress

                ' Append the cell address to the string builder
                stringBuilder.Append(cellAddress)

                ' Write the contents of the string builder to a text file named "output.txt"
                File.WriteAllText("output.txt", stringBuilder.ToString())
            End Using
            ExcelDocViewer("output.txt")
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
