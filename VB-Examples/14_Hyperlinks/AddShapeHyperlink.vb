Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace AddShapeHyperlink
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new workbook object
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\AddShapeHyperlink.xlsx")

			' Get the reference to the first sheet in the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Get all the shapes in the sheet
			Dim prstGeomShapeType As PrstGeomShapeCollection = sheet.PrstGeomShapes

			' Set the hyperlink for each shape
			For i As Integer = 0 To prstGeomShapeType.Count - 1
				' Get the shape
				Dim shape As XlsPrstGeomShape = CType(prstGeomShapeType(i), XlsPrstGeomShape)

				' Set the hyperlink address
				shape.HyLink.Address = "https://www.e-iceblue.com/Download/download-excel-for-net-now.html"
			Next i

			' Specify the filename for the resulting Excel file
			Dim result As String = "AddShapeHyperlink-out.xlsx"

			' Save the workbook to the specified file in Excel 2010 format
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose of the workbook object
			workbook.Dispose()

			' View the document using a file viewer
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
