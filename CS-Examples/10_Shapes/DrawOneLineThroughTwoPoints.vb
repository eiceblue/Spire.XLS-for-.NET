Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace DrawOneLineThroughTwoPoints
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Add a typed line shape to the worksheet and cast it to XlsLineShape
            Dim line1 As XlsLineShape = TryCast(worksheet.TypedLines.AddLine(), XlsLineShape)
            line1.LeftColumn = 3
            line1.TopRow = 3
            line1.LeftColumnOffset = 0
            line1.TopRowOffset = 0

            line1.RightColumn = 4
            line1.BottomRow = 5
            line1.RightColumnOffset = 0
            line1.BottomRowOffset = 0

            ' Add another typed line shape to the worksheet and set its start and end points
            Dim line2 As XlsLineShape = TryCast(worksheet.TypedLines.AddLine(), XlsLineShape)
            line2.StartPoint = New Point(30, 50)
            line2.EndPoint = New Point(20, 80)

            ' Specify the resulting file name for saving
            Dim result As String = "DrawOneLineThroughTwoPoints.xlsx"

            ' Save the workbook to a file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)
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
