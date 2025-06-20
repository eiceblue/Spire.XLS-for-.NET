Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace SetPivotFieldsConditionalFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")

            ' Get the worksheet named "PivotTable"
            Dim worksheet As Worksheet = workbook.Worksheets("PivotTable")

            ' Get the first PivotTable from the worksheet
            Dim table As PivotTable = CType(worksheet.PivotTables(0), PivotTable)

            ' Get the collection of PivotConditionalFormats of the PivotTable
            Dim pcfs As PivotConditionalFormatCollection = table.PivotConditionalFormats

            ' Add a new PivotConditionalFormat to the collection for the first data field of the PivotTable
            Dim pc As PivotConditionalFormat = pcfs.AddPivotConditionalFormat(table.DataFields(0))

            ' Add a new condition to the PivotConditionalFormat
            Dim cf As Spire.Xls.Core.IConditionalFormat = pc.AddCondition()

            ' Set the format type of the condition to NotContainsBlanks
            cf.FormatType = ConditionalFormatType.NotContainsBlanks

            ' Set the fill pattern of the condition to Solid
            cf.FillPattern = ExcelPatternType.Solid

            ' Set the background color of the condition to Yellow
            cf.BackColor = Color.Yellow

            ' Save the modified workbook to a new file named "output.xlsx" in Excel 2016 format
            workbook.SaveToFile("output.xlsx", ExcelVersion.Version2016)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the txt file
            ExcelDocViewer("output.xlsx")
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
