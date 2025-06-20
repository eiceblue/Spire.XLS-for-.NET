Imports Spire.Xls

Namespace AddImageHyperlink

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the width of column 0 to 22
            sheet.Columns(0).ColumnWidth = 22

            ' Set the text of cell A1 as "Image Hyperlink"
            sheet.Range("A1").Text = "Image Hyperlink"

            ' Set the vertical alignment of cell A1 to Top
            sheet.Range("A1").Style.VerticalAlignment = VerticalAlignType.Top

            ' Specify the path of the picture file
            Dim picPath As String = "..\..\..\..\..\..\Data\SpireXls.png"

            ' Add a picture to the worksheet at the specified position (row 2, column 1) with the specified picture file path
            Dim picture As ExcelPicture = sheet.Pictures.Add(2, 1, picPath)

            ' Set a hyperlink for the picture with the specified URL and open it in a new window
            picture.SetHyperLink("https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html", True)

            ' Specify the output file name
            Dim output As String = "AddImageHyperlink.xlsx"

            ' Save the modified workbook to a specified path with Excel 2013 format
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
