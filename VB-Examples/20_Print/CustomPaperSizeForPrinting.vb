Imports Spire.Xls


Namespace CustomPaperSizeForPrinting
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			'Load an excel file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CustomPaperSizeForPrinting.xlsx")

			Dim worksheet As Worksheet = workbook.Worksheets(0)
			'Set the paper size to the printer's custom paper size
			worksheet.PageSetup.CustomPaperSizeName = "customPaper"

			'Use the default printer to print
			workbook.PrintDocument.Print()

		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub

		Private Sub label1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles label1.Click

		End Sub
	End Class
End Namespace
