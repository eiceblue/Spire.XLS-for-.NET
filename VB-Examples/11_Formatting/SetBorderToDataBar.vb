Imports Spire.Xls
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet.Collections

Namespace SetBorderToDataBar
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()
            ' Load an existing Excel file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_9.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the conditional formats collection for the worksheet
            Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats(0)
            ' Get the first conditional format from the collection
            Dim cf As IConditionalFormat = xcfs(0)
            ' Get the DataBar object from the conditional format
            Dim dataBar1 As Spire.Xls.DataBar = cf.DataBar
            ' Set the type of border for the data bar (solid)
            dataBar1.BarBorder.Type = Spire.Xls.Core.Spreadsheet.ConditionalFormatting.DataBarBorderType.DataBarBorderSolid
            ' Set the color of the border to red
            dataBar1.BarBorder.Color = Color.Red

            ' Set the number value of cell E1 to 200
            sheet("E1").NumberValue = 200

            ' Add a new conditional formats collection to the worksheet
            Dim xcfs2 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
            ' Add the range E1 to the conditional formats collection
            xcfs2.AddRange(sheet.Range("E1"))
            ' Add a new conditional format to the collection
            Dim cf2 As IConditionalFormat = xcfs2.AddCondition()
            ' Set the format type to DataBar
            cf2.FormatType = ConditionalFormatType.DataBar
            ' Set the type of border for the data bar (solid)
            cf2.DataBar.BarBorder.Type = Spire.Xls.Core.Spreadsheet.ConditionalFormatting.DataBarBorderType.DataBarBorderSolid
            ' Set the color of the border to red
            cf2.DataBar.BarBorder.Color = Color.Red
            ' Set the color of the data bar to green-yellow
            cf2.DataBar.BarColor = Color.GreenYellow

            ' Save the workbook to a new file with the name "SetBorderToDataBar_result.xlsx"
            Dim result As String = "SetBorderToDataBar_result.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2013)

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
