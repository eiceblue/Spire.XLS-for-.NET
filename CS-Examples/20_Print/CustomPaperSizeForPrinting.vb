Imports Spire.Xls


Namespace CustomPaperSizeForPrinting
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "CustomPaperSizeForPrinting.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CustomPaperSizeForPrinting.xlsx")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Set the custom paper size name for the page setup of the worksheet
            worksheet.PageSetup.CustomPaperSizeName = "customPaper"

	    'Custom the paper size directly
	    'sheet.PageSetup.SetCustomPaperSize(224, CSng(50))

	    'Set the page orientation
	    'sheet.PageSetup.Orientation = PageOrientationType.Portrait

            ' Print the workbook using the default printer and settings
            workbook.PrintDocument.Print()

            ' Release the resources used by the workbook
            workbook.Dispose()

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
