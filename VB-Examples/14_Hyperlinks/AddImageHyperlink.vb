Imports Spire.Xls

Namespace AddImageHyperlink

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add the description text
			sheet.Columns(0).ColumnWidth = 22
			sheet.Range("A1").Text = "Image Hyperlink"
			sheet.Range("A1").Style.VerticalAlignment = VerticalAlignType.Top

			'Insert an image to a specific cell
			Dim picPath As String = "..\..\..\..\..\..\Data\SpireXls.png"
			Dim picture As ExcelPicture = sheet.Pictures.Add(2, 1, picPath)
			'Add a hyperlink to the image
			picture.SetHyperLink("https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html", True)

			'Save the document
			Dim output As String = "AddImageHyperlink.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

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
