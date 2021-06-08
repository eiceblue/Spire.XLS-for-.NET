Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes


Namespace SetShapeOrder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim wb As New Workbook()
			'Load an excel file
			wb.LoadFromFile("..\..\..\..\..\..\Data\SetShapeOrder.xlsx")

			'Bring the picture forward one level
			wb.Worksheets(0).Pictures(0).ChangeLayer(ShapeLayerChangeType.BringForward)

			'Bring the image in fron of all other objects
			wb.Worksheets(1).Pictures(0).ChangeLayer(ShapeLayerChangeType.BringToFront)

			'Send the shape back one level
			Dim shape As XlsShape = TryCast(wb.Worksheets(2).PrstGeomShapes(1), XlsShape)
			shape.ChangeLayer(ShapeLayerChangeType.SendBackward)

			'Send the shape behind all other objects
			shape = TryCast(wb.Worksheets(3).PrstGeomShapes(1), XlsShape)
			shape.ChangeLayer(ShapeLayerChangeType.SendToBack)

			Dim result As String = "SetShapeOrder_result.xlsx"
			'Save to file
			wb.SaveToFile(result, ExcelVersion.Version2010)
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
