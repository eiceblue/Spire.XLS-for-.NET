Imports Spire.Xls

Namespace SetTheme
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim srcWorkbook As New Workbook()
			'Load an excel file
			srcWorkbook.LoadFromFile("..\..\..\..\..\..\Data\SetTheme.xlsx")
			Dim srcWorksheet As Worksheet = srcWorkbook.Worksheets(0)

			Dim workbook As New Workbook()
			workbook.Worksheets.Clear()
			workbook.Worksheets.AddCopy(srcWorksheet)

			'1. Copy the theme of the workbook
			'workbook.CopyTheme(srcWorkbook);

			'2. Set a certain type of color of the default theme in the workbook
			workbook.SetThemeColor(ThemeColorType.Dk1, Color.SkyBlue)

			Dim result As String = "SetTheme_result.xlsx"
			'Save to file
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

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
