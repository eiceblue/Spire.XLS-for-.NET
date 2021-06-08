Imports Spire.Xls

Namespace SetPositionAndAlignment

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set two font styles which will be used in comments
			Dim font1 As ExcelFont = workbook.CreateFont()
			font1.FontName = "Calibri"
			font1.Color = Color.Firebrick
			font1.IsBold = True
			font1.Size = 12
			Dim font2 As ExcelFont = workbook.CreateFont()
			font2.FontName = "Calibri"
			font2.Color = Color.Blue
			font2.Size = 12
			font2.IsBold = True

			'Add comment 1 and set its size, text, position and alignment
			sheet.Range("G5").Text = "Spire.XLS"
			Dim Comment1 As ExcelComment = sheet.Range("G5").Comment
			Comment1.IsVisible = True
			Comment1.Height = 150
			Comment1.Width = 300
			Comment1.RichText.Text = "Spire.XLS for .Net:" & vbLf & "Standalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. "
			Comment1.RichText.SetFont(0, 19, font1)
			Comment1.TextRotation = TextRotationType.LeftToRight
			'Set the position of Comment
			Comment1.Top = 20
			Comment1.Left = 40
			'Set the alignment of text in Comment
			Comment1.VAlignment = CommentVAlignType.Center
			Comment1.HAlignment = CommentHAlignType.Justified

			'Add comment2 and set its size, text, position and alignment for comparison
			sheet.Range("D14").Text = "E-iceblue"
			Dim Comment2 As ExcelComment = sheet.Range("D14").Comment
			Comment2.IsVisible = True
			Comment2.Height = 150
			Comment2.Width = 300
			Comment2.RichText.Text = "About E-iceblue: " & vbLf & "We focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents."
			Comment2.TextRotation = TextRotationType.LeftToRight
			Comment2.RichText.SetFont(0, 16, font2)
			'Set the position of Comment
			Comment2.Top = 170
			Comment2.Left = 450
			'Set the alignment of text in Comment
			Comment2.VAlignment = CommentVAlignType.Top
			Comment2.HAlignment = CommentHAlignType.Justified

			'Save the document
			Dim output As String = "SetPositionAndAlignment.xlsx"
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
