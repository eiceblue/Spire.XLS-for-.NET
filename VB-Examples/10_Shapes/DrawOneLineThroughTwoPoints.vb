Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace DrawOneLineThroughTwoPoints
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'1)Draw a line according to relative position
			Dim line1 As XlsLineShape = TryCast(worksheet.TypedLines.AddLine(), XlsLineShape)
			line1.LeftColumn = 3
			line1.TopRow = 3
			line1.LeftColumnOffset = 0
			line1.TopRowOffset = 0

			line1.RightColumn = 4
			line1.BottomRow = 5
			line1.RightColumnOffset = 0
			line1.BottomRowOffset = 0

			'2)Draw a line according to absolute position(pixels).
			Dim line2 As XlsLineShape = TryCast(worksheet.TypedLines.AddLine(), XlsLineShape)
			line2.StartPoint = New Point(30, 50)
			line2.EndPoint = New Point(20, 80)

			'Save to file
			Dim result As String = "DrawOneLineThroughTwoPoints.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)

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
