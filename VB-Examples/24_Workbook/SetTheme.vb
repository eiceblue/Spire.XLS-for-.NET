Imports Spire.Xls

Namespace SetTheme
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to hold the source workbook
            Dim srcWorkbook As New Workbook()

            ' Load an existing Excel file into the source workbook
            srcWorkbook.LoadFromFile("..\..\..\..\..\..\Data\SetTheme.xlsx")

            ' Get the first worksheet from the source workbook
            Dim srcWorksheet As Worksheet = srcWorkbook.Worksheets(0)

            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Clear any existing worksheets in the workbook
            workbook.Worksheets.Clear()

            ' Add a copy of the source worksheet to the workbook
            workbook.Worksheets.AddCopy(srcWorksheet)

            ' Copy the theme from the source workbook to the current workbook (commented out)
            ' workbook.CopyTheme(srcWorkbook)

            ' Set the theme color for the workbook's theme
            workbook.SetThemeColor(ThemeColorType.Dk1, Color.SkyBlue)

            ' Specify the output file name
            Dim result As String = "SetTheme_result.xlsx"

            ' Save the workbook to the specified output file in Excel 2010 format
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

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
