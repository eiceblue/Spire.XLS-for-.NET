Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes


Namespace SetShapeOrder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim wb As New Workbook()

            ' Load an existing Excel file into the workbook
            wb.LoadFromFile("..\..\..\..\..\..\Data\SetShapeOrder.xlsx")

            ' Bring the first picture in the first worksheet forward by changing its layer
            wb.Worksheets(0).Pictures(0).ChangeLayer(ShapeLayerChangeType.BringForward)

            ' Bring the first picture in the second worksheet to the front by changing its layer
            wb.Worksheets(1).Pictures(0).ChangeLayer(ShapeLayerChangeType.BringToFront)

            ' Get the first PrstGeomShape in the third worksheet and cast it to XlsShape type
            Dim shape As XlsShape = TryCast(wb.Worksheets(2).PrstGeomShapes(1), XlsShape)
            ' Send the shape backward by changing its layer
            shape.ChangeLayer(ShapeLayerChangeType.SendBackward)

            ' Get the first PrstGeomShape in the fourth worksheet and cast it to XlsShape type
            shape = TryCast(wb.Worksheets(3).PrstGeomShapes(1), XlsShape)
            ' Send the shape to the back by changing its layer
            shape.ChangeLayer(ShapeLayerChangeType.SendToBack)

            ' Specify the file name for the resulting Excel file
            Dim result As String = "SetShapeOrder_result.xlsx"

            ' Save the workbook to the specified file in Excel 2010 format
            wb.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            wb.Dispose()
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
