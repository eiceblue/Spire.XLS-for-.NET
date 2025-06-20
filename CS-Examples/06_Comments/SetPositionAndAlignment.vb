Imports Spire.Xls

Namespace SetPositionAndAlignment

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create the first font object with specified properties
            Dim font1 As ExcelFont = workbook.CreateFont()
            font1.FontName = "Calibri"
            font1.Color = Color.Firebrick
            font1.IsBold = True
            font1.Size = 12

            ' Create the second font object with specified properties
            Dim font2 As ExcelFont = workbook.CreateFont()
            font2.FontName = "Calibri"
            font2.Color = Color.Blue
            font2.Size = 12
            font2.IsBold = True

            ' Set the text of cell G5 in the worksheet
            sheet.Range("G5").Text = "Spire.XLS"

            ' Get the comment object associated with cell G5
            Dim Comment1 As ExcelComment = sheet.Range("G5").Comment

            ' Make the comment visible
            Comment1.IsVisible = True

            ' Set the height and width of the comment box
            Comment1.Height = 150
            Comment1.Width = 300

            ' Set the text of the comment
            Comment1.RichText.Text = "Spire.XLS :" & vbLf & "Standalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc."

            ' Apply font1 to the first 20 characters of the comment text
            Comment1.RichText.SetFont(0, 19, font1)

            ' Set the text rotation of the comment
            Comment1.TextRotation = TextRotationType.LeftToRight

            ' Set the top and left position of the comment box
            Comment1.Top = 20
            Comment1.Left = 40

            ' Set the vertical and horizontal alignment of the text within the comment box
            Comment1.VAlignment = CommentVAlignType.Center
            Comment1.HAlignment = CommentHAlignType.Justified

            ' Set the text of cell D14 in the worksheet
            sheet.Range("D14").Text = "E-iceblue"

            ' Get the comment object associated with cell D14
            Dim Comment2 As ExcelComment = sheet.Range("D14").Comment

            ' Make the comment visible
            Comment2.IsVisible = True

            ' Set the height and width of the comment box
            Comment2.Height = 150
            Comment2.Width = 300

            ' Set the text of the comment
            Comment2.RichText.Text = "About E-iceblue: " & vbLf & "We focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents."

            ' Set the text rotation of the comment
            Comment2.TextRotation = TextRotationType.LeftToRight

            ' Apply font2 to the first 16 characters of the comment text
            Comment2.RichText.SetFont(0, 16, font2)

            ' Set the top and left position of the comment box
            Comment2.Top = 170
            Comment2.Left = 450

            ' Set the vertical and horizontal alignment of the text within the comment box
            Comment2.VAlignment = CommentVAlignType.Top
            Comment2.HAlignment = CommentHAlignType.Justified

            ' Specify the output file name
            Dim output As String = "SetPositionAndAlignment.xlsx"

            ' Save the workbook to the specified output file in Excel 2013 format
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
