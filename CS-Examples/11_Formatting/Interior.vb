Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace Interior
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

            ' Set the Excel version for the workbook to 2007
            workbook.Version = ExcelVersion.Version2007

            ' Determine the maximum number of colors available in the palette
            Dim maxColor As Integer = System.Enum.GetValues(GetType(ExcelColors)).Length

            ' Initialize a random number generator with a seed value of 10000000
            Dim random As New Random(10000000)

            ' Loop through rows 2 to 39
            For i As Integer = 2 To 39
                ' Generate a random known color for the cell background
                Dim backKnownColor As ExcelColors = CType(random.Next(1, maxColor \ 2), ExcelColors)

                ' Set the header labels for columns A to D
                sheet.Range("A1").Text = "Color Name"
                sheet.Range("B1").Text = "Red"
                sheet.Range("C1").Text = "Green"
                sheet.Range("D1").Text = "Blue"

                ' Merge cells E1 to K1 and set the text as "Gradient"
                sheet.Range("E1:K1").Merge()
                sheet.Range("E1:K1").Text = "Gradient"

                ' Apply formatting to the header row
                sheet.Range("A1:K1").Style.Font.IsBold = True
                sheet.Range("A1:K1").Style.Font.Size = 11

                ' Get the name of the backKnownColor
                Dim colorName As String = backKnownColor.ToString()

                ' Set the color name in column A
                sheet.Range(String.Format("A{0}", i)).Text = colorName

                ' Set the Red component value in column B
                sheet.Range(String.Format("B{0}", i)).NumberValue = workbook.GetPaletteColor(backKnownColor).R

                ' Set the Green component value in column C
                sheet.Range(String.Format("C{0}", i)).NumberValue = workbook.GetPaletteColor(backKnownColor).G

                ' Set the Blue component value in column D
                sheet.Range(String.Format("D{0}", i)).NumberValue = workbook.GetPaletteColor(backKnownColor).B

                ' Merge cells Ei to Ki and set the text as the color name
                sheet.Range(String.Format("E{0}:K{0}", i)).Merge()
                sheet.Range(String.Format("E{0}:K{0}", i)).Text = colorName

                ' Apply gradient formatting to cells Ei to Ki
                sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.FillPattern = ExcelPatternType.Gradient

                ' Set the background known color for the gradient
                sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.Gradient.BackKnownColor = backKnownColor

                ' Set the foreground color as white for the gradient
                sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.Gradient.ForeKnownColor = ExcelColors.White

                ' Set the gradient style to vertical
                sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.Gradient.GradientStyle = GradientStyleType.Vertical

                ' Set the gradient variant to ShadingVariants1
                sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.Gradient.GradientVariant = GradientVariantsType.ShadingVariants1
            Next i

            ' Auto-fit the width of column 1 (column A)
            sheet.AutoFitColumn(1)

            ' Specify the filename to save the modified workbook
            Dim result As String = "Result-Interior.xlsx"

            ' Save the workbook to the specified file path using Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer(result)
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
