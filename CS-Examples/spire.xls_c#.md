# spire.xls c# excel Chart DataTable
## Enable data table for chart in Excel 
//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
//Get the first chart
Chart chart = sheet.Charts[0];
// Enable the data table for the chart
chart.HasDataTable = true;

# spire.xls c# excel Chart Shapes
## Add picture to chart in Excel file
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
// Get the first chart
Chart chart = sheet.Charts[0];
// Add the picture in chart
chart.Shapes.AddPicture(@".\SpireXls.png");

# spire.xls c# excel Chart Shapes
## Add textbox to Excel chart
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
// Get the first chart
Chart chart = sheet.Charts[0];
// Add a Textbox
var textbox = chart.Shapes.AddTextBox();
// Set the width of the textbox
textbox.Width = 1200;
// Set the height of the textbox
textbox.Height = 320;
// Set the top position of the textbox
textbox.Top = 480;
// Set the left position of the textbox
textbox.Left = 1000;
textbox.Text = "This is a textbox";

# spire.xls c# excel Chart TrendLine
## Add trendlines to Excel charts
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
//select chart and set logarithmic trendline
Chart chart = sheet.Charts[0];
chart.ChartTitle = "Logarithmic Trendline";
chart.Series[0].TrendLines.Add(TrendLineType.Logarithmic);
//select chart and set moving_average trendline
Chart chart1 = sheet.Charts[1];
chart1.ChartTitle = "Moving Average Trendline";
chart1.Series[0].TrendLines.Add(TrendLineType.Moving_Average);
//select chart and set linear trendline
Chart chart2 = sheet.Charts[2];
chart2.ChartTitle = "Linear Trendline";
chart2.Series[0].TrendLines.Add(TrendLineType.Linear);
//select chart and set exponential trendline
Chart chart3 = sheet.Charts[3];
chart3.ChartTitle = "Exponential Trendline";
chart3.Series[0].TrendLines.Add(TrendLineType.Exponential);

# spire.xls c# excel Chart Series
## Adjust bar space (GapWidth, Overlap) in Excel charts
//Get the first worksheet from workbook and then get the first chart from the worksheet
Worksheet ws = workbook.Worksheets[0];
Chart chart = ws.Charts[0];
//Adjust the space between bars
foreach (var cs in chart.Series)
{
    cs.Format.Options.GapWidth = 200;
    cs.Format.Options.Overlap = 0;
}

# spire.xls c# excel Chart ChartArea
## Apply soft edge effect to Excel chart area
//Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
//Get the chart
Chart chart = sheet.Charts[0];
//Specify the size of the soft edge. Value can be set from 0 to 100
chart.ChartArea.Shadow.SoftEdge = 25;

# spire.xls c# excel Chart Position
## Modify chart size (Width, Height) and position (LeftColumn, TopRow) in Excel
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
//Get the chart
Chart chart = sheet.Charts[0];
//Change chart size
chart.Width = 600;
chart.Height = 500;
//Change chart position
chart.LeftColumn = 3;
chart.TopRow = 7;

# spire.xls c# excel Chart DataLabels
## Modify chart data label text for a specific datapoint
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
//Get the chart
Chart chart = sheet.Charts[0];
//Change data label of the frist datapoint of the first series
chart.Series[0].DataPoints[0].DataLabels.Text = "changed data label";

# spire.xls c# excel Chart DataRange
## Modify chart data range in Excel workbook
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Get chart
Chart chart = sheet.Charts[0];
// Change data range
chart.DataRange = sheet.Range["A1:C4"];

# spire.xls c# excel Chart Axis
## Change chart major gridlines color
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Get the chart
Chart chart = sheet.Charts[0];
//Change the color of major gridlines
chart.PrimaryValueAxis.MajorGridLines.LineProperties.Color = System.Drawing.Color.Red;

# spire.xls c# excel Chart Series
## Change chart series fill color
//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
//Get the first chart
Chart chart = sheet.Charts[0];
//Get the second series
var cs = chart.Series[1];
//Set the fill type
cs.Format.Fill.FillType = ShapeFillType.SolidColor;
//Change the fill color
cs.Format.Fill.ForeColor = System.Drawing.Color.Orange;

# spire.xls c# excel Chart Axis
## Set chart axis titles and font size
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
//Get the chart
Chart chart = sheet.Charts[0];
//Set axis title
chart.PrimaryCategoryAxis.Title = "Category Axis";
chart.PrimaryValueAxis.Title = "Value axis";
//Set font size
chart.PrimaryCategoryAxis.Font.Size = 12;
chart.PrimaryValueAxis.Font.Size = 12;

# spire.xls c# excel Chart Conversion
## Convert Excel chart to PNG image
//Save chart as image
var stream = workbook.SaveChartAsImage(workbook.Worksheets[0], 0);

# spire.xls c# excel Chart Type
## Create Box and Whisker Chart in Excel
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Add a new chart
Chart officeChart = sheet.Charts.Add();
//Set the chart title
officeChart.ChartTitle = "Yearly Vehicle Sales";
// Set chart type as Box and Whisker
officeChart.ChartType = ExcelChartType.BoxAndWhisker;
// Set data range in the worksheet
officeChart.DataRange = sheet["A1:E17"];
// Box and Whisker settings on first series
var seriesA = officeChart.Series[0];
seriesA.DataFormat.ShowInnerPoints = false;
seriesA.DataFormat.ShowOutlierPoints = true;
seriesA.DataFormat.ShowMeanMarkers = true;
seriesA.DataFormat.ShowMeanLine = false;
seriesA.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian;
// Box and Whisker settings on second series
var seriesB = officeChart.Series[1];
seriesB.DataFormat.ShowInnerPoints = false;
seriesB.DataFormat.ShowOutlierPoints = true;
seriesB.DataFormat.ShowMeanMarkers = true;
seriesB.DataFormat.ShowMeanLine = false;
seriesB.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.InclusiveMedian;
// Box and Whisker settings on third series
var seriesC = officeChart.Series[2];
seriesC.DataFormat.ShowInnerPoints = false;
seriesC.DataFormat.ShowOutlierPoints = true;
seriesC.DataFormat.ShowMeanMarkers = true;
seriesC.DataFormat.ShowMeanLine = false;
seriesC.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian;

# spire.xls c# excel Chart Type
## Add and save a bubble chart in Excel
// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];
// Add a Bubble chart to the worksheet
Chart chart = sheet.Charts.Add(ExcelChartType.Bubble);
// Set the title of the chart
chart.ChartTitle = "Bubble";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
// Specify the range of data for the chart
chart.DataRange = sheet.Range["A1:C5"];
chart.SeriesDataFromRange = false;
// Set the range of values for the bubbles in the chart
chart.Series[0].Bubbles = sheet.Range["C2:C5"];
// Set the position of the chart on the worksheet
chart.LeftColumn = 7;
chart.TopRow = 6;
chart.RightColumn = 16;
chart.BottomRow = 29;

# spire.xls c# excel Chart PivotTable
## Create chart from pivot table in Excel
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets[0];
// Get the pivot table
var pt = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Add a chart based on the pivot table to the second worksheet
workbook.Worksheets[1].Charts.Add(ExcelChartType.BarClustered, pt);

# spire.xls c# excel Chart Type
## Create doughnut chart in Excel
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];
// Insert data into the worksheet
sheet.Range["A1"].Value = "Country";
sheet.Range["A1"].Style.Font.IsBold = true;
sheet.Range["A2"].Value = "Cuba"; // ... (rest of data omitted for brevity)
sheet.Range["B5"].NumberValue = 8500;
// Add a new chart and set its type to Doughnut
Chart chart = sheet.Charts.Add();
chart.ChartType = ExcelChartType.Doughnut;
// Set the data range for the chart
chart.DataRange = sheet.Range["A1:B5"];
chart.SeriesDataFromRange = false;
// Set the position of the chart on the worksheet
chart.LeftColumn = 4; // ... (positioning omitted for brevity)
chart.BottomRow = 22;
// Set the chart title
chart.ChartTitle = "Market share by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
// Enable percentage labels for each data point
foreach (var cs in chart.Series)
{
    cs.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = true;
}
// Set the legend position to the top
chart.Legend.Position = LegendPositionType.Top;

# spire.xls c# excel Chart Type
## Create funnel chart from Excel file
//Find the first worksheet
var sheet = workbook.Worksheets[0];
//Add a new chart
var officeChart = sheet.Charts.Add();
//Set chart type as Funnel
officeChart.ChartType = ExcelChartType.Funnel;
//Set data range in the worksheet
officeChart.DataRange = sheet.Range["A1:B6"];
//Set the chart title
officeChart.ChartTitle = "Funnel";
//Formatting the legend and data label option
officeChart.HasLegend = false;
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;

# spire.xls c# excel Chart Type
## Create histogram chart in Excel
//Find the first worksheet
var sheet = workbook.Worksheets[0];
//Add a new chart
var officeChart = sheet.Charts.Add();
//Set chart type as histogram
officeChart.ChartType = ExcelChartType.Histogram;
//Set data range in the worksheet
officeChart.DataRange = sheet["A1:A15"];
officeChart.TopRow = 1; // ... (positioning omitted for brevity)
officeChart.RightColumn = 12;
//Category axis bin settings
officeChart.PrimaryCategoryAxis.BinWidth = 8;
//Gap width settings
officeChart.Series[0].DataFormat.Options.GapWidth = 6;
//Set the chart title and axis title
officeChart.ChartTitle = "Height Data";
officeChart.PrimaryValueAxis.Title = "Number of students";
officeChart.PrimaryCategoryAxis.Title = "Height";
//Hiding the legend
officeChart.HasLegend = false;

# spire.xls c# excel Chart Type
## Create multi-level clustered bar chart in Excel
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Write data to cells
sheet.Range["A1"].Text = "Main Category"; // ... (data population omitted for brevity)
sheet.Range["C9"].Value = "57";
// Vertically merge cells
sheet.Range["A2:A5"].Merge();
sheet.Range["A6:A9"].Merge();
sheet.AutoFitColumn(1);
sheet.AutoFitColumn(2);
// Add a clustered bar chart to worksheet
Chart chart = sheet.Charts.Add(ExcelChartType.BarClustered);
chart.ChartTitle = "Value";
chart.PlotArea.Fill.FillType = ShapeFillType.NoFill;
chart.Legend.Delete();
chart.LeftColumn = 5; // ... (positioning omitted for brevity)
chart.RightColumn = 14;
// Set the data source of series data
chart.DataRange = sheet.Range["C2:C9"];
chart.SeriesDataFromRange = false;
// Set the data source of category labels
var serie = chart.Series[0];
serie.CategoryLabels = sheet.Range["A2:B9"];
// Show multi-level category labels
chart.PrimaryCategoryAxis.MultiLevelLable = true;

# spire.xls c# excel Chart Type
## Create and view a Pareto chart from Excel data.
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Add chart
Chart officeChart = sheet.Charts.Add();
// Set chart type as Pareto
officeChart.ChartType = ExcelChartType.Pareto;
//Set data range in the worksheet
officeChart.DataRange = sheet["A2:B8"];
officeChart.TopRow = 1; // ... (positioning omitted for brevity)
officeChart.RightColumn = 12;
officeChart.PrimaryCategoryAxis.IsBinningByCategory = true;
officeChart.PrimaryCategoryAxis.OverflowBinValue = 5;
officeChart.PrimaryCategoryAxis.UnderflowBinValue = 1;
// Formatting Pareto line
officeChart.Series[0].ParetoLineFormat.LineProperties.Color = System.Drawing.Color.Blue;
// Gap width settings
officeChart.Series[0].DataFormat.Options.GapWidth = 6;
// Set the chart title
officeChart.ChartTitle = "Expenses";
// Hiding the legend
officeChart.HasLegend = false;

# spire.xls c# excel Chart PivotTable
## Create pivot chart from Excel data
//get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
//get the first pivot table in the worksheet
var pivotTable = sheet.PivotTables[0];
//create a clustered column chart based on the pivot table
Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, pivotTable);
//set chart position
chart.TopRow = 12; // ... (positioning omitted for brevity)
chart.BottomRow = 30;
chart.ChartTitle = "Product";
chart.PrimaryCategoryAxis.MultiLevelLable = true;

# spire.xls c# excel Chart Type
## Generate radar chart from Excel data
//Create an empty worksheet
workbook.CreateEmptySheets(1);
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
sheet.Name = "Chart data";
sheet.GridLinesVisible = false;
//Add a new chart
Chart chart = sheet.Charts.Add();
//Set position of chart
chart.LeftColumn = 1; // ... (positioning omitted for brevity)
chart.BottomRow = 29;
//Set region of chart data
chart.DataRange = sheet.Range["A1:C5"];
chart.SeriesDataFromRange = false;
// Set chart type
chart.ChartType = ExcelChartType.Radar; // or RadarFilled based on condition
//Set chart title
chart.ChartTitle = "Sale market by region";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
chart.PlotArea.Fill.Visible = false;
chart.Legend.Position = LegendPositionType.Corner;

# spire.xls c# excel Chart Type
## Create and view a sunburst chart in Excel
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Add chart
Chart officeChart = sheet.Charts.Add();
// Set chart type as Sunburst
officeChart.ChartType = ExcelChartType.SunBurst;
//Set data range in the worksheet
officeChart.DataRange = sheet["A1:D16"];
officeChart.TopRow = 1; // ... (positioning omitted for brevity)
officeChart.RightColumn = 14;
// Set the chart title
officeChart.ChartTitle = "Sales by quarter";
// Formatting data labels
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
// Hiding the legend
officeChart.HasLegend = false;

# spire.xls c# excel Chart Type
## Create TreeMap chart from Excel data
//Find the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Add chart
Chart officeChart = sheet.Charts.Add();
// Set chart type as TreeMap
officeChart.ChartType = ExcelChartType.TreeMap;
// Set data range in the worksheet
officeChart.DataRange = sheet["A2:C11"];
officeChart.TopRow = 1; // ... (positioning omitted for brevity)
officeChart.RightColumn = 14;
// Set the chart title
officeChart.ChartTitle = "Area by countries";
// Set the Treemap label option
officeChart.Series[0].DataFormat.TreeMapLabelOption = ExcelTreeMapLabelOption.Banner;
// Formatting data labels
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;

# spire.xls c# excel Chart Type
## Create waterfall chart from Excel workbook.
// Get the first worksheet
var sheet = workbook.Worksheets[0];
// Add a new chart to the worksheet
var officeChart = sheet.Charts.Add();
// Set chart type as waterfall
officeChart.ChartType = ExcelChartType.WaterFall;
// Set data range for the chart from the worksheet
officeChart.DataRange = sheet["A2:B8"];
// Set chart position and size
officeChart.TopRow = 1; // ... (positioning omitted for brevity)
officeChart.RightColumn = 12;
// Set certain data points in the chart as totals
officeChart.Series[0].DataPoints[3].SetAsTotal = true;
officeChart.Series[0].DataPoints[6].SetAsTotal = true;
// Show connector lines between data points
officeChart.Series[0].Format.ShowConnectorLines = true;
// Set the chart title
officeChart.ChartTitle = "Waterfall Chart";
// Format data labels and legend options
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
officeChart.Legend.Position = LegendPositionType.Right;

# spire.xls c# excel Chart Series
## Create scatter chart with custom markers (color, style, size, transparency)
// Create an empty sheet
workbook.CreateEmptySheets(1);
Worksheet sheet = workbook.Worksheets[0];
// Add some sample data
sheet.Name = "Demo";
sheet.Range["A1"].Value = "Tom"; // ... (data population omitted for brevity)
sheet.Range["B7"].NumberValue = 4.7;
//Create a Scatter-Markers chart based on the sample data
Chart chart = sheet.Charts.Add(ExcelChartType.ScatterMarkers);
chart.DataRange = sheet.Range["A1:B7"];
chart.PlotArea.Visible = false;
chart.SeriesDataFromRange = false;
chart.TopRow = 5; // ... (positioning omitted for brevity)
chart.RightColumn = 11;
chart.ChartTitle = "Chart with Markers";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 10;
//Format the markers in the chart by setting the background color, foreground color, type, size and transparency
Spire.Xls.Charts.ChartSerie cs1 = chart.Series[0];
cs1.DataFormat.MarkerBackgroundColor = System.Drawing.Color.RoyalBlue;
cs1.DataFormat.MarkerForegroundColor = System.Drawing.Color.WhiteSmoke;
cs1.DataFormat.MarkerSize = 7;
cs1.DataFormat.MarkerStyle = ChartMarkerType.PlusSign;
cs1.DataFormat.MarkerTransparencyValue = 0.8;
Spire.Xls.Charts.ChartSerie cs2 = chart.Series[1];
cs2.DataFormat.MarkerBackgroundColor = System.Drawing.Color.Pink;
cs2.DataFormat.MarkerSize = 9;
cs2.DataFormat.MarkerStyle = ChartMarkerType.Triangle;
cs2.DataFormat.MarkerTransparencyValue = 0.9;

# spire.xls c# excel Chart DataLabels
## Customize chart data labels (Callout, CategoryName, SeriesName, LegendKey)
//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
//Get the first chart
Chart chart = sheet.Charts[0];
// Enable data labels and customize callout settings for each series in the chart
foreach (Spire.Xls.Charts.ChartSerie cs in chart.Series)
{
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasWedgeCallout = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = true;
}

# spire.xls c# excel Chart Legend
## Remove first two legends entries from Excel chart
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Get the chart
Chart chart = sheet.Charts[0];
// Delete the first and the second legend entries from the chart
chart.Legend.LegendEntries[0].Delete();
chart.Legend.LegendEntries[1].Delete();

# spire.xls c# excel Chart Series
## Create discontinuous chart from Excel data (non-contiguous ranges)
// Get the first sheet
Worksheet sheet = book.Worksheets[0];
// Add a chart
Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered);
chart.SeriesDataFromRange = false;
// Set the position of chart
chart.LeftColumn = 1; // ... (positioning omitted for brevity)
chart.BottomRow = 24;
// Add a series
Spire.Xls.Charts.ChartSerie cs1 = (Spire.Xls.Charts.ChartSerie)chart.Series.Add();
// Set the name of the cs1
cs1.Name = sheet.Range["B1"].Value;
// Set discontinuous values for cs1
cs1.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"]);
cs1.Values = sheet.Range["B2:B3"].AddCombinedRange(sheet.Range["B5:B6"]).AddCombinedRange(sheet.Range["B8:B9"]);
//Set the chart type
cs1.SerieType = ExcelChartType.ColumnClustered;
// Add another series (cs2) similarly...
Spire.Xls.Charts.ChartSerie cs2 = (Spire.Xls.Charts.ChartSerie)chart.Series.Add();
cs2.Name = sheet.Range["C1"].Value;
cs2.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"]);
cs2.Values = sheet.Range["C2:C3"].AddCombinedRange(sheet.Range["C5:C6"]).AddCombinedRange(sheet.Range["C8:C9"]);
cs2.SerieType = ExcelChartType.ColumnClustered;
// Set the chart title
chart.ChartTitle = "Chart";
chart.ChartTitleArea.Size = 20;
chart.ChartTitleArea.Color = System.Drawing.Color.Black;
// Disable major grid lines on the primary value axis
chart.PrimaryValueAxis.HasMajorGridLines = false;

# spire.xls c# excel Chart Series
## Add a new series to an existing line chart in Excel
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Get the line chart
Chart chart = sheet.Charts[0];
// Add a new series
Spire.Xls.Charts.ChartSerie cs = chart.Series.Add("Added");
// Set the values for the series
cs.Values = sheet.Range["I1:L1"];

# spire.xls c# excel Chart Type
## Create an exploded doughnut chart in Excel.
// Get the first sheet and set its name
Worksheet sheet = workbook.Worksheets[0];
sheet.Name = "ExplodedDoughnut";
// Add a chart
Chart chart = sheet.Charts.Add();
chart.ChartType = ExcelChartType.DoughnutExploded;
// Set position of chart
chart.LeftColumn = 1; // ... (positioning omitted for brevity)
chart.BottomRow = 29;
// Set region of chart data
chart.DataRange = sheet.Range["A1:B5"];
chart.SeriesDataFromRange = false;
// Chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
foreach (Spire.Xls.Charts.ChartSerie cs in chart.Series)
{
    cs.Format.Options.IsVaryColor = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
}
chart.PlotArea.Fill.Visible = false;
chart.Legend.Position = LegendPositionType.Top;

# spire.xls c# excel Chart TrendLine
## Extract trendline equation from Excel chart and save to text file.
// Get the chart from the first worksheet
Chart chart = workbook.Worksheets[0].Charts[0];
// Get the trendline of the chart and then extract the equation of the trendline
var trendLine = chart.Series[1].TrendLines[0];
string formula = trendLine.Formula;
StringBuilder sb = new StringBuilder();
sb.AppendLine("The equation is: " + formula);
// Save to Text file
string output = "ExtractTrendline.txt";
File.WriteAllText(output, sb.ToString());

# spire.xls c# excel Chart Fill
## Fill chart area and plot area with image in Excel
//Get the first worksheet from workbook
Worksheet ws = workbook.Worksheets[0];
//Get the first chart
Chart chart = ws.Charts[0];
// A. Fill chart area with image
chart.ChartArea.Fill.CustomPicture((@".\background.png"));
chart.PlotArea.Fill.Transparency = 0.9;

# spire.xls c# excel Chart Series
## Fill chart series markers with picture, texture, or pattern
// Specify the path of the input Excel file.
// string inputFile = @".\FillChartMarker.xlsx"; // This line might be context for the user.
string imageFile = @".\E-iceblueLogo.png";
Worksheet worksheet = workbook.Worksheets[0];
Chart chart = worksheet.Charts[0];
// Series 1: Fill marker with picture
chart.Series[0].Format.LineProperties.Color = System.Drawing.Color.Yellow;
chart.Series[0].Format.MarkerStyle = ChartMarkerType.Picture;
var markerFill1 = chart.Series[0].DataFormat.MarkerFill;
markerFill1.CustomPicture(imageFile);
// Series 2: Fill marker with texture
var markerFill2 = chart.Series[1].DataFormat.MarkerFill;
chart.Series[1].Format.LineProperties.Color = System.Drawing.Color.Red;
markerFill2.Texture = GradientTextureType.Granite;
// Series 3: Fill marker with pattern
chart.Series[2].Format.LineProperties.Color = System.Drawing.Color.Blue;
var markerFill3 = chart.Series[2].DataFormat.MarkerFill;
markerFill3.Pattern = GradientPatternType.Pat10Percent;
markerFill3.ForeColor = System.Drawing.Color.LightGray;
markerFill3.BackColor = System.Drawing.Color.Orange;

# spire.xls c# excel Chart Axis
## Format chart axis (MajorUnit, MinorUnit, MaxValue, MinValue, NumberFormat, etc.)
Worksheet sheet = workbook.Worksheets[0];
sheet.Name = "FormatAxis";
//Add a chart
Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered);
chart.DataRange = sheet.Range["B1:B9"];
chart.SeriesDataFromRange = false;
chart.PlotArea.Visible = false;
chart.TopRow = 10; // ... (positioning omitted for brevity)
chart.RightColumn = 10;
chart.ChartTitle = "Chart with Customized Axis";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
Spire.Xls.Charts.ChartSerie cs1 = chart.Series[0];
cs1.CategoryLabels = sheet.Range["A2:A9"];
//Format axis
chart.PrimaryValueAxis.MajorUnit = 8;
chart.PrimaryValueAxis.MinorUnit = 2;
chart.PrimaryValueAxis.MaxValue = 50;
chart.PrimaryValueAxis.MinValue = 0;
chart.PrimaryValueAxis.IsReverseOrder = false;
chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkOutside;
chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkInside;
chart.PrimaryValueAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionNextToAxis;
chart.PrimaryValueAxis.CrossesAt = 0;
//Set NumberFormat
chart.PrimaryValueAxis.NumberFormat = "$#,##0";
chart.PrimaryValueAxis.IsSourceLinked = false;
// Format Series DataPoints
foreach (Spire.Xls.Charts.ChartDataPoint dataPoint in cs1.DataPoints)
{
    dataPoint.DataFormat.Fill.FillType = ShapeFillType.SolidColor;
    dataPoint.DataFormat.Fill.ForeColor = System.Drawing.Color.LightGreen;
    dataPoint.DataFormat.Fill.Transparency = 0.3;
}

# spire.xls c# excel Chart Type
## Create doughnut and pie charts (gauge chart) in Excel workbook
Worksheet sheet = workbook.Worksheets[0];
sheet.Name = "Gauge Chart";
// Add a Doughnut chart
Chart chart = sheet.Charts.Add(ExcelChartType.Doughnut);
chart.DataRange = sheet.Range["A1:A5"];
chart.SeriesDataFromRange = false;
chart.HasLegend = true;
// Set the position of chart
chart.LeftColumn = 2; // ... (positioning omitted for brevity)
chart.BottomRow = 25;
// Get the series 1 (Doughnut)
Spire.Xls.Charts.ChartSerie cs1 = (Spire.Xls.Charts.ChartSerie)chart.Series["Value"];
cs1.Format.Options.DoughnutHoleSize = 60;
cs1.DataFormat.Options.FirstSliceAngle = 270;
// Set the fill color for Doughnut slices
cs1.DataPoints[0].DataFormat.Fill.ForeColor = System.Drawing.Color.Yellow;
cs1.DataPoints[1].DataFormat.Fill.ForeColor = System.Drawing.Color.PaleVioletRed;
cs1.DataPoints[2].DataFormat.Fill.ForeColor = System.Drawing.Color.DarkViolet;
cs1.DataPoints[3].DataFormat.Fill.Visible = false; // Hide last slice for gauge effect
// Add a series with pie chart (Pointer)
Spire.Xls.Charts.ChartSerie cs2 = (Spire.Xls.Charts.ChartSerie)chart.Series.Add("Pointer", ExcelChartType.Pie);
// Set the value for Pie (Pointer)
cs2.Values = sheet.Range["D2:D4"];
cs2.UsePrimaryAxis = false;
cs2.DataPoints[0].DataLabels.HasValue = true;
cs2.DataFormat.Options.FirstSliceAngle = 270;
cs2.DataPoints[0].DataFormat.Fill.Visible = false; // Pointer base
cs2.DataPoints[1].DataFormat.Fill.FillType = ShapeFillType.SolidColor; // Pointer needle
cs2.DataPoints[1].DataFormat.Fill.ForeColor = System.Drawing.Color.Black;
cs2.DataPoints[2].DataFormat.Fill.Visible = false; // Pointer tail

# spire.xls c# excel Chart Axis
## Extract category labels from Excel chart
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Get the first chart
Chart chart = sheet.Charts[0];
//Get the cell range of the category labels
CellRange cr = chart.PrimaryCategoryAxis.CategoryLabels;
StringBuilder sb = new StringBuilder();
foreach (var cell in cr)
{
    sb.Append(cell.Value + "\r\n");
}
//Save the result file
string result = "result.txt";
File.WriteAllText(result, sb.ToString());

# spire.xls c# excel Chart DataPoints
## Extract chart data point values and their addresses
// Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
// Get the chart
Chart chart = sheet.Charts[0];
// Get the first series of the chart
Spire.Xls.Charts.ChartSerie cs = chart.Series[0];
StringBuilder sb = new StringBuilder();
foreach (CellRange cell_range in cs.Values) // Corrected variable name
{
    sb.Append(cell_range.RangeAddress + "\r\n");
    //Get the data point value
    sb.Append("The value of the data point is " + cell_range.Value + "\r\n");
}
string result = "result.txt";
// Save the file
File.WriteAllText(result, sb.ToString());

# spire.xls c# excel Chart Worksheet
## Extract chart's parent worksheet name and save to text file.
//Access first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];
//Access the first chart inside this worksheet
Chart chart = worksheet.Charts[0];
//Get its worksheet
Worksheet wSheet = chart.Worksheet as Worksheet;
//Create StringBuilder to save
StringBuilder content = new StringBuilder();
//Set string format for displaying
string result_text = string.Format("Sheet Name: " + worksheet.Name + "\r\nCharts' sheet Name: " + wSheet.Name);
content.AppendLine(result_text);
//String for output file
String outputFile = "Output.txt";
//Save them to a txt file
File.WriteAllText(outputFile, content.ToString());

# spire.xls c# excel Chart Axis
## Hide Excel chart major gridlines
Worksheet sheet = workbook.Worksheets[0];
//Get the chart
Chart chart = sheet.Charts[0];
//Hide major gridlines
chart.PrimaryValueAxis.HasMajorGridLines = false;

# spire.xls c# excel Chart Type
## Create 3D or 2D line chart in Excel
Worksheet sheet = workbook.Worksheets[0];
sheet.Name = "Line Chart";
// Add a chart
Chart chart = sheet.Charts.Add();
bool use3D = true; // Example: set to true for 3D
// Set chart type based on input
if (use3D)
{
    chart.ChartType = ExcelChartType.Line3D;
}
else
{
    chart.ChartType = ExcelChartType.Line;
}
// Set region of chart data
chart.DataRange = sheet.Range["A1:E5"];
// Set position of the chart
chart.LeftColumn = 1; // ... (positioning omitted for brevity)
chart.BottomRow = 29;
// Set chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
// Customize axes and series
chart.PrimaryCategoryAxis.Title = "Month"; // ... (further customization omitted)
chart.PrimaryValueAxis.Title = "Sales (in Dollars)"; // ...
foreach (Spire.Xls.Charts.ChartSerie cs in chart.Series)
{
    cs.Format.Options.IsVaryColor = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
    if (!use3D) cs.DataFormat.MarkerStyle = ChartMarkerType.Circle;
}
chart.PlotArea.Fill.Visible = false;
chart.Legend.Position = LegendPositionType.Top;

# spire.xls c# excel Chart Series
## Add drop lines to Excel line chart series
// Get the first sheet
Worksheet worksheet = workbook.Worksheets[0];
// Get the first chart
Chart chart = worksheet.Charts[0];
// Add a drop lines to the first series
chart.Series[0].HasDroplines = true;

# spire.xls c# excel Chart Type
## Create 3D pie chart and save Excel file
Worksheet sheet = workbook.Worksheets[0];
sheet.Name = "Pie Chart";
//Add a chart
Chart chart = sheet.Charts.Add(ExcelChartType.Pie3D); // Or Pie for 2D
//Set region of chart data
chart.DataRange = sheet.Range["B2:B5"];
chart.SeriesDataFromRange = false;
//Set position of chart
chart.LeftColumn = 1; // ... (positioning omitted for brevity)
chart.BottomRow = 25;
//Chart title
chart.ChartTitle = "Sales by year";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
Spire.Xls.Charts.ChartSerie cs = chart.Series[0];
cs.CategoryLabels = sheet.Range["A2:A5"];
cs.Values = sheet.Range["B2:B5"];
cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
chart.PlotArea.Fill.Visible = false;

# spire.xls c# excel Chart Type
## Create and display a 3D pyramid chart in Excel.
Worksheet sheet = workbook.Worksheets[0];
sheet.Name = "Chart";
// Add a chart
Chart chart = sheet.Charts.Add();
// Set region of chart data
chart.DataRange = sheet.Range["B2:B5"];
chart.SeriesDataFromRange = false;
// Set position of the chart
chart.LeftColumn = 1; // ... (positioning omitted for brevity)
chart.BottomRow = 29;
chart.ChartType = ExcelChartType.PyramidClustered;
// Set chart title and axes
chart.ChartTitle = "Sales by year"; // ... (title/axis formatting omitted)
Spire.Xls.Charts.ChartSerie cs_pyramid = chart.Series[0];
cs_pyramid.CategoryLabels = sheet.Range["A2:A5"];
cs_pyramid.Format.Options.IsVaryColor = true;
chart.Legend.Position = LegendPositionType.Top;

# spire.xls c# excel Chart Remove
## Remove chart from Excel worksheet
//Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];
//Get the first chart from the first worksheet
Spire.Xls.Core.IChartShape chartShape = sheet.Charts[0];
//Remove the chart
chartShape.Remove();

# spire.xls c# excel Chart Position
## Resize and move chart in Excel file
//Get the chart from the first worksheet
Worksheet sheet = workbook.Worksheets[0];
Chart chart = sheet.Charts[0];
//Set position of the chart
chart.LeftColumn = 5;
chart.TopRow = 1;
//Resize the chart
chart.Width = 500;
chart.Height = 350;

# spire.xls c# excel Chart DataLabels
## Enhance Excel chart data labels with rich text and styles.
//Get first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];
//Get the first chart inside this worksheet
Chart chart = worksheet.Charts[0];
//Get the first datalabel of the first series
Spire.Xls.Charts.ChartDataLabels datalabel = chart.Series[0].DataPoints[0].DataLabels;
//Set the text
datalabel.Text = "Rich Text Label";
//Show the value
chart.Series[0].DataPoints[0].DataLabels.HasValue = true;
//Set styles for the text
chart.Series[0].DataPoints[0].DataLabels.Color = System.Drawing.Color.Red;
chart.Series[0].DataPoints[0].DataLabels.IsBold = true;

# spire.xls c# excel Chart 3D Format
## Rotate 3D chart (Elevation, Rotation) in Excel file
//Get the chart from the first worksheet
Worksheet sheet = workbook.Worksheets[0];
Chart chart = sheet.Charts[0];
//X rotation:
chart.Rotation = 30;
//Y rotation:
chart.Elevation = 20;

# spire.xls c# excel Chart DataLabels
## Create line chart with data labels (HasValue, HasLegendKey, Position, Color, FontName, Size etc.)
Worksheet sheet = workbook.Worksheets[0];
// Set sheet name and populate data
sheet.Name = "Demo";
sheet.Range["A1"].Value = "Month"; // ... (data population omitted for brevity)
sheet.Range["B7"].NumberValue = 28;
// Add a line chart with markers
Chart chart = sheet.Charts.Add(ExcelChartType.LineMarkers);
// Set chart data range and position
chart.DataRange = sheet.Range["B1:B7"];
chart.PlotArea.Visible = false; // ... (more settings omitted)
chart.RightColumn = 11;
// Set chart title
chart.ChartTitle = "Data Labels Demo"; // ...
Spire.Xls.Charts.ChartSerie cs1_label = chart.Series[0];
cs1_label.CategoryLabels = sheet.Range["A2:A7"];
// Customize data label settings
cs1_label.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
cs1_label.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = false;
cs1_label.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = false;
cs1_label.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = true;
cs1_label.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = true;
cs1_label.DataPoints.DefaultDataPoint.DataLabels.Delimiter = ". ";
cs1_label.DataPoints.DefaultDataPoint.DataLabels.Size = 9;
cs1_label.DataPoints.DefaultDataPoint.DataLabels.Color = System.Drawing.Color.Red;
cs1_label.DataPoints.DefaultDataPoint.DataLabels.FontName = "Calibri";
cs1_label.DataPoints.DefaultDataPoint.DataLabels.Position = DataLabelPositionType.Center;

# spire.xls c# excel Chart Series
## Set border color and style (CustomLineWeight) for chart series line
//Get the first worksheet from workbook and then get the first chart from the worksheet
Worksheet ws = workbook.Worksheets[0];
Chart chart = ws.Charts[0];
//Set CustomLineWeight property for Series line
(chart.Series[0].DataPoints[0].DataFormat.LineProperties as Spire.Xls.Core.Spreadsheet.Charts.XlsChartBorder).CustomLineWeight = 2.5f;
//Set color property for Series line
(chart.Series[0].DataPoints[0].DataFormat.LineProperties as Spire.Xls.Core.Spreadsheet.Charts.XlsChartBorder).Color = System.Drawing.Color.Red;

# spire.xls c# excel Chart Series
## Adjust Excel chart marker border widths (MarkerBorderWidth)
// Get the chart from the first worksheet
Chart chart = workbook.Worksheets[0].Charts[0];
// Set marker border width for series 1
chart.Series[0].DataFormat.MarkerBorderWidth = 1.5;
// Set marker border width for series 2
chart.Series[1].DataFormat.MarkerBorderWidth = 2.5;

# spire.xls c# excel Chart ChartArea
## Set Excel chart background color (ChartArea.ForeGroundColor)
//Get the first worksheet from workbook and then get the first chart from the worksheet
Worksheet ws = workbook.Worksheets[0];
Chart chart = ws.Charts[0];
//Set background color
chart.ChartArea.ForeGroundColor = System.Drawing.Color.LightYellow;

# spire.xls c# excel Chart Font
## Set font for legend and data label in Excel chart
//Get the first worksheet from workbook
Worksheet ws = workbook.Worksheets[0];
Chart chart = ws.Charts[0];
//Create a font with specified size and color
ExcelFont font = workbook.CreateFont();
font.Size = 14.0;
font.Color = System.Drawing.Color.Red;
//Apply the font to chart Legend
chart.Legend.TextArea.SetFont(font);
//Apply the font to chart DataLabel
foreach (Spire.Xls.Charts.ChartSerie cs_font in chart.Series)
{
    cs_font.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font);
}

# spire.xls c# excel Chart Font
## Set font for chart title and axis in Excel
//Set font for chart title and chart axis
Worksheet worksheet = workbook.Worksheets[0];
Chart chart = worksheet.Charts[0];
//Format the font for the chart title
chart.ChartTitleArea.Color = System.Drawing.Color.Blue;
chart.ChartTitleArea.Size = 20.0;
chart.ChartTitleArea.FontName = "Arial";
//Format the font for the chart Axis
chart.PrimaryValueAxis.Font.Color = System.Drawing.Color.Gold;
chart.PrimaryValueAxis.Font.Size = 10.0;
chart.PrimaryCategoryAxis.Font.FontName = "Arial";
chart.PrimaryCategoryAxis.Font.Color = System.Drawing.Color.Red;
chart.PrimaryCategoryAxis.Font.Size = 20.0;

# spire.xls c# excel Chart Legend
## Set chart legend background color to sky blue
// Get the first worksheet
Worksheet ws = workbook.Worksheets[0];
// Get the chart from the worksheet
Chart chart = ws.Charts[0];
// Access the legend frame format and set the background color
var legendFrame = chart.Legend.FrameFormat as Spire.Xls.Core.Spreadsheet.Charts.XlsChartFrameFormat;
legendFrame.Fill.FillType = ShapeFillType.SolidColor;
legendFrame.ForeGroundColor = System.Drawing.Color.SkyBlue;

# spire.xls c# excel Chart TrendLine
## Set trendline number format (DataLabel.NumberFormat) in Excel workbook
//Get the chart from the first worksheet
Chart chart = workbook.Worksheets[0].Charts[0];
//Get the trendline of the chart
Spire.Xls.Core.IChartTrendLine trendLine = chart.Series[1].TrendLines[0];
//Set the number format of trendLine to "#,##0.00"
trendLine.DataLabel.NumberFormat = "#,##0.00";

# spire.xls c# excel Chart DataLabels
## Create a bar chart with leader lines (DataLabels.ShowLeaderLines) in Excel
workbook.Version = ExcelVersion.Version2013;
//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
//Set value of specified range
sheet.Range["A1"].Value = "1"; // ... (data population omitted for brevity)
sheet.Range["C3"].Value = "9";
// Add a stacked bar chart
Chart chart = sheet.Charts.Add(ExcelChartType.BarStacked);
chart.DataRange = sheet.Range["A1:C3"];
chart.TopRow = 4; // ... (positioning omitted for brevity)
chart.Height = 300;
// Enable data labels with leader lines for each series
foreach (Spire.Xls.Charts.ChartSerie cs_leader in chart.Series)
{
    cs_leader.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
    cs_leader.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = true;
}

# spire.xls c# excel Chart Sparkline
## Generate sparklines in Excel workbook.
//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
//Add sparkline
SparklineGroup sparklineGroup = sheet.SparklineGroups.AddGroup(SparklineType.Line);
SparklineCollection sparklines = sparklineGroup.Add();
sparklines.Add(sheet["A2:D2"], sheet["E2"]);
sparklines.Add(sheet["A3:D3"], sheet["E3"]); // ... (more sparklines added similarly)
sparklines.Add(sheet["A11:D11"], sheet["E11"]);

# spire.xls c# excel Chart Conversion
## Convert Excel chart to SVG
//Get the second chartsheet by name
ChartSheet cs_svg = workbook.GetChartSheetByName("Chart1");
//Save to SVG stream
string output_svg = "ToSVG.svg";
FileStream fs_svg = new FileStream(output_svg, FileMode.Create);
cs_svg.ToSVGStream(fs_svg);
fs_svg.Flush();
fs_svg.Close();

# spire.xls c# excel Chart Font
## Embed custom font in Excel chart
//Get the first sheet
Worksheet sheet = workbook.Worksheets[0];
//Get the first chart
Chart chart = sheet.Charts[0];
//Load the font file from disk
workbook.CustomFontFilePaths = new string[] { @".\PT_Serif-Caption-Web-Regular.ttf" };
System.Collections.Hashtable fontResult = workbook.GetCustomFontParsedResult();
ArrayList valueList = new ArrayList(fontResult.Values);
//Apply the font for PrimaryValueAxis of chart
chart.PrimaryValueAxis.Font.FontName = valueList[0] as string;
//Apply the font for PrimaryCategoryAxis of chart
chart.PrimaryCategoryAxis.Font.FontName = valueList[0] as string;
//Apply the font for the first Spire.Xls.Charts.ChartSerie of chart
Spire.Xls.Charts.ChartSerie chartSerie1 = chart.Series[0];
chartSerie1.DataPoints.DefaultDataPoint.DataLabels.FontName = valueList[0] as string;

# spire.xls c# excel Chart Worksheet
## Move chart sheets within an Excel workbook
//Move the first chartsheet to the position of the third sheet(including chartsheet and worksheet)
workbook.Chartsheets[0].MoveSheet(2);

# spire.xls c# excel Pivot Tables
## Add Label and Value Filters to PivotTable Row Fields
//Retrieve the first pivot table from the second sheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = workbook.Worksheets[1].PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
//Add a label filter to the first row field of the pivot table
pt.RowFields[0].AddLabelFilter(PivotLabelFilterType.Between, "Argentina", "Nicaragua");
// Add a value filter on the first row field of the pivot table
pt.RowFields[0].AddValueFilter(PivotValueFilterType.LessThan, pt.DataFields[0], 5300000, null);
pt.CalculateData();

# spire.xls c# excel Pivot Tables
## Change PivotTable Data Source
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Define the range of cells to be used as the new data source
CellRange range = sheet.Range["A1:C15"];
// Get the first pivot table from the second worksheet
PivotTable table = workbook.Worksheets[1].PivotTables[0] as PivotTable;
// Change the data source of the pivot table to the new range
table.ChangeDataSource(range);
// Disable automatic refresh of the pivot table cache on load
table.Cache.IsRefreshOnLoad = false;

# spire.xls c# excel Pivot Tables
## Clear All Data Fields from a PivotTable
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];
// Get the first pivot table from the sheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Clear all the data fields in the pivot table
pt.DataFields.Clear();
// Calculate the pivot table data
pt.CalculateData();

# spire.xls c# excel Pivot Tables
## Set Consolidation Functions for PivotTable Data Fields
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Apply Average consolidation function to first data field
pt.DataFields[0].Subtotal = SubtotalTypes.Average;
// Apply Max consolidation function to second data field
pt.DataFields[1].Subtotal = SubtotalTypes.Max;
// Calculate data
pt.CalculateData();

# spire.xls c# excel Pivot Tables
## Create a New PivotTable
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Set the value to the cells
sheet.Range["A1"].Value = "Product";
sheet.Range["B1"].Value = "Month";
sheet.Range["C1"].Value = "Count";
sheet.Range["A2"].Value = "SpireDoc";
sheet.Range["A3"].Value = "SpireDoc";
sheet.Range["A4"].Value = "SpireXls";
sheet.Range["C2"].Value = "10";
sheet.Range["C3"].Value = "15";
sheet.Range["C4"].Value = "9";
// Add a PivotTable to the worksheet
CellRange dataRange = sheet.Range["A1:C4"]; // Simplified data for example
PivotCache cache = workbook.PivotCaches.Add(dataRange);
PivotTable pt = sheet.PivotTables.Add("NewPivotTable", sheet.Range["E10"], cache);
// Drag the fields to the row area.
PivotField pf = pt.PivotFields["Product"] as PivotField;
pf.Axis = AxisTypes.Row;
// Drag the field to the data area.
pt.DataFields.Add(pt.PivotFields["Count"], "SUM of Count", SubtotalTypes.Sum);
// Set PivotTable style
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12;
pt.CalculateData();

# spire.xls c# excel Pivot Tables
## Set Custom Names for PivotTable Fields
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];
// Access the first pivot table in the worksheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTable = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Set a custom name for the row field
pivotTable.RowFields[0].CustomName = "custom_rowName";
// Set a custom name for the column field
pivotTable.ColumnFields[0].CustomName = "custom_colName";
// Set a custom name for the data field
pivotTable.DataFields[0].CustomName = "custom_DataName";
// Calculate the pivot table data
pivotTable.CalculateData();

# spire.xls c# excel Pivot Tables
## Disable PivotTable Ribbon/Wizard
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];
// Get the first pivot table from the sheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
//Disable ribbon for this pivot table
pt.EnableWizard = false;

# spire.xls c# excel Pivot Tables
## Expand or Collapse Rows in a PivotTable
// Get the first worksheet.
Worksheet sheet = workbook.Worksheets[0];
// Get the first pivot table from the sheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTable = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Calculate data.
pivotTable.CalculateData();
// Collapse the rows for "Vendor No" 3501.
(pivotTable.PivotFields["Vendor No"] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3501", true);
// Expand the rows for "Vendor No" 3502.
(pivotTable.PivotFields["Vendor No"] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3502", false);

# spire.xls c# excel Pivot Tables
## Format PivotTable DataField to Show as Percentage of Column
Worksheet sheet = workbook.Worksheets[0];
// Get the first pivot table from the sheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Access the data field.
PivotDataField pivotDataField = pt.DataFields[0];
// Set data display format to PercentageOfColumn
pivotDataField.ShowDataAs = PivotFieldFormatType.PercentageOfColumn;
pt.CalculateData(); // Recalculate to apply changes

# spire.xls c# excel Pivot Tables
## Set PivotTable Built-in Style and Layout Options
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];
// Get the first pivot table from the worksheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTable = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Set the built-in style for the pivot table appearance
pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleLight10;
// Enable the display of grid drop zone in the pivot table
pivotTable.Options.ShowGridDropZone = true;
// Set the row layout type to compact in the pivot table
pivotTable.Options.RowLayout = PivotTableLayoutType.Compact;

# spire.xls c# excel Pivot Tables
## Get PivotTable Refreshed Information (User and Date)
// Get first worksheet of the workbook
Worksheet worksheet = workbook.Worksheets[0];
// Get the first pivot table
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTable = worksheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Get the refreshed information
System.DateTime dateTime = pivotTable.Cache.RefreshDate;
string refreshedBy = pivotTable.Cache.RefreshedBy;
// Output information (e.g., to console or a file)
System.Console.WriteLine("Pivot table refreshed by: " + refreshedBy);
System.Console.WriteLine("Pivot table refreshed date: " + dateTime.ToString());

# spire.xls c# excel Pivot Tables
## Group PivotTable Data by Date Range
// Get the first worksheet in the workbook
Worksheet sheet = workbook.Worksheets[0];
// Get the first pivot table in the worksheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Get the first row field in the pivot table
Spire.Xls.Core.IPivotField field = pt.RowFields[0];
// Set the start and end dates for grouping
System.DateTime start = new System.DateTime(2023, 1, 5);
System.DateTime end = new System.DateTime(2023, 3, 2);
// Set the group by type to days
PivotGroupByTypes[] types = new PivotGroupByTypes[] { PivotGroupByTypes.Days };
// Create a new group with the specified start and end dates, group by type, and interval
field.CreateGroup(start, end, types, 10);
// Calculate the pivot table data
pt.CalculateData();
// Refresh the pivot table cache
pt.Cache.IsRefreshOnLoad = true;

# spire.xls c# excel Pivot Tables
## Set PivotTable Report Layout to Tabular
// Get the first worksheet
Worksheet worksheet = workbook.Worksheets[0];
// Get the first PivotTable
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable xlsPivotTable = (Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable)worksheet.PivotTables[0];
// Set the PivotTable layout type to Tabular
xlsPivotTable.Options.ReportLayout = PivotTableLayoutType.Tabular;

# spire.xls c# excel Pivot Tables
## Refresh PivotTable Data After Source Update
// Get the data source worksheet.
Worksheet dataSourceSheet = workbook.Worksheets[1]; // Assuming data is on the second sheet
// Update the data source of PivotTable.
dataSourceSheet.Range["D2"].Value = "999"; // Example: Change a value in the source
// Get the PivotTable worksheet.
Worksheet pivotSheet = workbook.Worksheets[0]; // Assuming pivot table is on the first sheet
// Get the PivotTable that was built on the data source.
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = pivotSheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Refresh the data of PivotTable.
pt.Cache.IsRefreshOnLoad = true; // Ensures it refreshes if opened by Excel
pt.CalculateData(); // Explicitly recalculate/refresh

# spire.xls c# excel Pivot Tables
## Repeat All Item Labels in a PivotTable
// Iterate through each pivot table in the "Pivot" worksheet
foreach (Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt in workbook.Worksheets["Pivot"].PivotTables)
{
    // Set the RepeatAllItemLabels property to true for the pivot table
    pt.Options.RepeatAllItemLabels = true;
    // Calculate the data for the pivot table
    pt.CalculateData();
    // Refresh the cache for the pivot table
    pt.Cache.IsRefreshOnLoad = true;
}

# spire.xls c# excel Pivot Tables
## Repeat Item Labels for Specific Fields in PivotTable
// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
// Add an empty worksheet
Worksheet sheet2 = workbook.CreateEmptySheet();
sheet2.Name = "Pivot Table";
// Define the data range for the pivot table
CellRange dataRange = sheet.Range["A1:D9"];
// Create a pivot cache using the data range
PivotCache cache = workbook.PivotCaches.Add(dataRange);
// Add a pivot table to the pivot sheet using the pivot cache
PivotTable pt = sheet2.PivotTables.Add("NewPivotTable", sheet.Range["A1"], cache);
// Set the VendorNo field as a row field
PivotField r1 = pt.PivotFields["VendorNo"] as PivotField;
r1.Axis = AxisTypes.Row;
pt.Options.RowHeaderCaption = "VendorNo";
r1.Subtotals = SubtotalTypes.None;
// Enable repeating item labels for the VendorNo field
r1.RepeatItemLabels = true;
// Set the row layout type to tabular
pt.Options.RowLayout = PivotTableLayoutType.Tabular;
// Add data field
pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.None);

# spire.xls c# excel Pivot Tables
## Set Various Formatting Options for PivotTable
// Get the sheet where the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];
// Access the first pivot table in the sheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Enable automatic formatting for the pivot table report
pt.Options.IsAutoFormat = true;
// Show grand totals for rows in the pivot table report
pt.ShowRowGrand = true;
// Show grand totals for columns in the pivot table report
pt.ShowColumnGrand = true;
// Display a custom string in cells that contain null values
pt.DisplayNullString = true;
pt.NullString = "null";
// Set the layout of the pivot table report for page fields
pt.PageFieldOrder = PagesOrderType.DownThenOver;

# spire.xls c# excel Pivot Tables
## Set Formatting for a Specific PivotTable Field
// Get the sheet in which the pivot table is located
Worksheet sheet = workbook.Worksheets["PivotTable"];
// Access the first pivot table in the worksheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTable = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Access the first pivot field in the pivot table
PivotField pivotField = pivotTable.PivotFields[0] as PivotField;
// Set the sort type of the pivot field to ascending
pivotField.SortType = PivotFieldSortType.Ascending;
// Enable displaying subtotals at the top of groups for the pivot field
pivotField.SubtotalTop = true;
// Set the subtotal type of the pivot field to Count
pivotField.Subtotals = SubtotalTypes.Count;
// Enable auto show for the pivot field
pivotField.IsAutoShow = true;

# spire.xls c# excel Pivot Tables
## Set Conditional Formatting for PivotTable Fields
// Get the worksheet with the PivotTable
Worksheet worksheet = workbook.Worksheets["PivotTable"];
// Get the PivotTable from the worksheet
PivotTable table = (PivotTable)worksheet.PivotTables[0];
// Add a conditional format to the PivotTable
PivotConditionalFormatCollection pcfs = table.PivotConditionalFormats;
// Apply to the first data field
PivotConditionalFormat pc = pcfs.AddPivotConditionalFormat(table.DataFields[0]);
Spire.Xls.Core.IConditionalFormat cf = pc.AddCondition();
// Example: Highlight non-blank cells
cf.FormatType = ConditionalFormatType.NotContainsBlanks;
cf.FillPattern = ExcelPatternType.Solid;
cf.BackColor = System.Drawing.Color.Yellow;

# spire.xls c# excel Pivot Tables
## Show PivotTable Data Field in Row Area
// Get the worksheet where the pivot table is located
Worksheet sheet = workbook.Worksheets[1]; // Assuming PivotTable is on the second sheet
// Access the pivot table in the worksheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTable = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Show the data field in the row area of the pivot table
pivotTable.ShowDataFieldInRow = true;
// Calculate the data in the pivot table
pivotTable.CalculateData();

# spire.xls c# excel Pivot Tables
## Show Subtotals in PivotTable
// Get the worksheet that contains the pivot table
Worksheet sheet = workbook.Worksheets["Pivot Table"];
// Get the first pivot table from the worksheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Enable the display of subtotals in the pivot table
pt.ShowSubtotals = true;

# spire.xls c# excel Pivot Tables
## Sort a PivotTable Field
// Get the first worksheet from the workbook
Worksheet sheet = workbook.Worksheets[0];
// Add an empty worksheet to the workbook and set its name
Worksheet sheet2 = workbook.CreateEmptySheet();
sheet2.Name = "Pivot Table";
// Specify the data source range for the pivot table
CellRange dataRange = sheet.Range["A1:C9"];
// Create a pivot cache using the data range
PivotCache cache = workbook.PivotCaches.Add(dataRange);
// Add a pivot table to the second worksheet using the specified cache
PivotTable pt = sheet2.PivotTables.Add("NewPivotTable", sheet.Range["A1"], cache);
// Configure the pivot table settings
PivotField r1 = pt.PivotFields["No"] as PivotField;
r1.Axis = AxisTypes.Row;
// Sort the "No" field in descending order
r1.SortType = PivotFieldSortType.Descending;
// Add a data field to the pivot table
pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.None);
pt.CalculateData();

# spire.xls c# excel Pivot Tables
## Update PivotTable Data Source and Refresh
// Access the "Data" worksheet
Worksheet dataSheet = workbook.Worksheets["Data"];
// Modify the data source by changing the value in cell A2 to "NewValue"
dataSheet.Range["A2"].Text = "NewValue";
// Modify the data source by changing the value in cell D2 to 28000
dataSheet.Range["D2"].NumberValue = 28000;
// Access the worksheet containing the pivot table
Worksheet pivotSheet = workbook.Worksheets["PivotTable"];
// Get the first pivot table from the worksheet
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pt = pivotSheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
// Set the pivot table's cache to refresh on load
pt.Cache.IsRefreshOnLoad = true;
// Calculate and update the pivot table data
pt.CalculateData();

# spire.xls c# excel Worksheets
## Set Worksheet Name
sheet.Name = "MyNewSheetName";

# spire.xls c# excel Worksheets
## Add Label Shape to Worksheet
Spire.Xls.Core.ILabelShape label = sheet.LabelShapes.AddLabel(10, 2, 30, 200);
label.Text = "This is a Label";

# spire.xls c# excel Worksheets
## Add ListBox Control to Worksheet
//Assume data for ListBox is in A7:A12
sheet.Range["A7"].Text = "Item1";
sheet.Range["A8"].Text = "Item2";
Spire.Xls.Core.IListBox listBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80);
listBox.ListFillRange = sheet.Range["A7:A12"];
listBox.SelectionType = SelectionType.Single;

# spire.xls c# excel Worksheets
## Add ScrollBar Shape to Worksheet
//Assume cell B10 exists for linking
sheet.Range["B10"].Value2 = 1;
Spire.Xls.Core.IScrollBarShape scrollBar = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20);
scrollBar.LinkedCell = sheet.Range["B10"];
scrollBar.Min = 1;
scrollBar.Max = 150;

# spire.xls c# excel Worksheets
## Create Table (ListObject) in Worksheet
//Assume sheet.LastRow and sheet.LastColumn are valid
Spire.Xls.Core.IListObject table = sheet.ListObjects.Create("MyTable", sheet.Range[1, 1, sheet.LastRow, sheet.LastColumn]);
table.BuiltInTableStyle = Spire.Xls.TableBuiltInStyles.TableStyleMedium9;

# spire.xls c# excel Worksheets
## Set PageSetup SummaryRowBelow Property
sheet.PageSetup.IsSummaryRowBelow = false; // True to display summary rows below detail, False for above

# spire.xls c# excel Worksheets
## Group Rows in Worksheet
//Group rows from 2 to 9, do not collapse them by default
sheet.GroupByRows(2, 9, false);

# spire.xls c# excel Worksheets
## Export Worksheet Data to DataTable
System.Data.DataTable dt = sheet.ExportDataTable();

# spire.xls c# excel Worksheets
## Export Specific Range of Worksheet to DataTable with Options
//Assume sheet.LastDataRow and sheet.LastDataColumn are valid
Spire.Xls.ExportTableOptions options = new Spire.Xls.ExportTableOptions();
options.KeepDataFormat = false;
options.RenameStrategy = Spire.Xls.RenameStrategy.Digit;
System.Data.DataTable table = sheet.ExportDataTable(1, 1, sheet.LastDataRow, sheet.LastDataColumn, options);

# spire.xls c# excel Worksheets
## Insert DataTable into Worksheet
//Assume 'System.Data.DataTable dataTable' exists and is populated
sheet.InsertDataTable(dataTable, true, 1, 1, -1, -1); // Insert with column headers starting at A1

# spire.xls c# excel Worksheets
## Find All Strings in Worksheet
Spire.Xls.CellRange[] foundRanges = sheet.FindAllString("searchText", false, false); // (text, caseSensitive, matchWholeWord)
if (foundRanges != null)
{
    foreach (Spire.Xls.CellRange range in foundRanges)
    {
        // Process found range, e.g., range.Style.Color = System.Drawing.Color.Yellow;
        Console.WriteLine("Found at: " + range.RangeAddress);
    }
}

# spire.xls c# excel Worksheets
## Find All Numbers in Worksheet
Spire.Xls.CellRange[] foundRanges = sheet.FindAllNumber(100.5, true); // (number, searchIntegersOnly)
if (foundRanges != null)
{
    foreach (Spire.Xls.CellRange range in foundRanges)
    {
        // Process found range
        Console.WriteLine("Found number at: " + range.RangeAddress);
    }
}

# spire.xls c# excel Worksheets
## Insert ArrayList into Worksheet
System.Collections.ArrayList dataList = new System.Collections.ArrayList();
dataList.Add("Item 1");
dataList.Add(123);
dataList.Add(System.DateTime.Now);
//Insert dataList into worksheet starting at cell A1, as a vertical list (isVertical = true)
sheet.InsertArrayList(dataList, 1, 1, true);

# spire.xls c# excel Worksheets
## Insert DataColumns into Worksheet
System.Data.DataTable dataTable = new System.Data.DataTable("SampleData");
dataTable.Columns.Add("ID", typeof(int));
dataTable.Columns.Add("Name", typeof(string));
dataTable.Rows.Add(1, "Product A");
dataTable.Rows.Add(2, "Product B");
System.Data.DataColumn[] columnsToInsert = new System.Data.DataColumn[] { dataTable.Columns["ID"], dataTable.Columns["Name"] };
//Insert specified columns into worksheet starting at A1, with column headers
sheet.InsertDataColumns(columnsToInsert, true, 1, 1);

# spire.xls c# excel Worksheets
## Insert DataView into Worksheet
System.Data.DataTable dataTable = new System.Data.DataTable("SourceTable");
dataTable.Columns.Add("City", typeof(string));
dataTable.Columns.Add("Population", typeof(int));
dataTable.Rows.Add("New York", 8000000);
dataTable.Rows.Add("London", 9000000);
dataTable.Rows.Add("Paris", 2000000);
dataTable.DefaultView.Sort = "Population DESC"; // Example: sort data
//Insert DataView into worksheet starting at A1, with column headers
sheet.InsertDataView(dataTable.DefaultView, true, 1, 1);

# spire.xls c# excel Worksheets
## Add TextBox to Worksheet
Spire.Xls.Core.ITextBoxShape textbox = sheet.TextBoxes.AddTextBox(5, 2, 30, 150); // (topRow, leftColumn, height, width)
textbox.Text = "Sample Textbox Content";

# spire.xls c# excel Worksheets
## Add CheckBox to Worksheet
Spire.Xls.Core.ICheckBox checkBox = sheet.CheckBoxes.AddCheckBox(7, 2, 20, 100); // (topRow, leftColumn, height, width)
checkBox.Text = "Option 1";
checkBox.CheckState = CheckState.Checked;

# spire.xls c# excel Worksheets
## Add RadioButton to Worksheet
Spire.Xls.Core.IRadioButton radioButton = sheet.RadioButtons.Add(9, 2, 20, 100); // (topRow, leftColumn, height, width)
radioButton.Text = "Select Me";
radioButton.CheckState = CheckState.Checked;

# spire.xls c# excel Worksheets
## Add ComboBox to Worksheet
//Assume data for ComboBox is in A20:A22
sheet.Range["A20"].Text = "Choice A";
sheet.Range["A21"].Text = "Choice B";
Spire.Xls.Core.IComboBoxShape comboBox = sheet.ComboBoxes.AddComboBox(11, 2, 20, 120) as Spire.Xls.Core.IComboBoxShape; // (topRow, leftColumn, height, width)
comboBox.ListFillRange = sheet.Range["A20:A22"];
comboBox.SelectedIndex = 0;

# spire.xls c# excel Worksheets
## Replace All Text in Worksheet with Style
//Assume 'CellStyle oldStyle' and 'CellStyle newStyle' are defined
//This will replace all occurrences of "OldText" with "NewText" and apply newStyle if oldStyle matches
//To replace text irrespective of style, pass null for oldStyle
//sheet.ReplaceAll("OldText", null, "NewText", newStyle);
//To replace text only if it has a specific style:
Spire.Xls.CellStyle styleToFind = sheet.Workbook.Styles.Add("StyleToFind"); // Example
styleToFind.Font.Color = System.Drawing.Color.Red;
sheet.Range["A1"].Text = "TextToReplace";
sheet.Range["A1"].Style = styleToFind;

Spire.Xls.CellStyle replacementStyle = sheet.Workbook.Styles.Add("ReplacementStyle"); // Example
replacementStyle.Font.Color = System.Drawing.Color.Green;

sheet.ReplaceAll("TextToReplace", styleToFind, "ReplacedText", replacementStyle);

# spire.xls c# excel Worksheets
## Copy CellRange from This Worksheet to Another Worksheet
var sourceSheet = workbook.Worksheets[0];
var destinationSheet = workbook.Worksheets[1];
Spire.Xls.CellRange rangeToCopy = sourceSheet.Range["A1:B5"]; // Example
//Copy rangeToCopy from sourceSheet to destinationSheet, starting at row 2, column 1 (A2)
//with style and updating references
sourceSheet.Copy(rangeToCopy, destinationSheet, 2, 1, true, true);

# spire.xls c# excel Worksheets
## Insert Array into Worksheet
object[,] dataArray = new object[,] {
    { "Name", "Age" },
    { "Alice", 30 },
    { "Bob", 24 }
};
//Insert dataArray into worksheet starting at cell A1
sheet.InsertArray(dataArray, 1, 1);

# spire.xls c# excel Worksheets
## Apply Subtotal to Range in Worksheet
//Assume data in range A1:B10 with categories in column A and values in column B
//sheet.Range["A1"].Value = "Category"; sheet.Range["B1"].Value = "Amount"; ...
Spire.Xls.CellRange dataRange = sheet.Range["A1:B10"]; // Example range
//Apply subtotal on change in the first column (index 0), summing values in the second column (index 1)
sheet.Subtotal(dataRange, 0, new int[] { 1 }, Spire.Xls.SubtotalTypes.Sum, true, false, true);

# spire.xls c# excel Worksheets
## Apply Style to Used Cells in Worksheet
Spire.Xls.CellStyle newStyle = sheet.Workbook.Styles.Add("MyUsedCellStyle");
newStyle.Font.Color = System.Drawing.Color.Blue;
newStyle.Interior.Color = System.Drawing.Color.LightYellow;
//Apply the style to all cells that contain data or formatting
sheet.ApplyStyle(newStyle, false, false); // (style, applyFont, applyBorder) - customize as needed

# spire.xls c# excel Worksheets
## Get Count of All Cells in Worksheet
int cellCount = sheet.Cells.Length;
Console.WriteLine("Total cells in worksheet: " + cellCount);

# spire.xls c# excel Worksheets
## Get Merged Cell Ranges in Worksheet
Spire.Xls.CellRange[] mergedCells = sheet.MergedCells;
if (mergedCells != null)
{
    foreach (Spire.Xls.CellRange mergedRange in mergedCells)
    {
        Console.WriteLine("Merged range: " + mergedRange.RangeAddress);
        // To unmerge: mergedRange.UnMerge();
    }
}

# spire.xls c# excel Worksheets
## Add AutoFilter and Apply Fill Color Filter
//Set the range for AutoFilter
sheet.AutoFilters.Range = sheet.Range["A1:C10"];
//Get the first filter column (column A)
Spire.Xls.Core.Spreadsheet.AutoFilter.FilterColumn filterColumn = (Spire.Xls.Core.Spreadsheet.AutoFilter.FilterColumn)sheet.AutoFilters[0];
//Add a fill color filter for Red color
sheet.AutoFilters.AddFillColorFilter(filterColumn, System.Drawing.Color.Red);
//Apply the filter
sheet.AutoFilters.Filter();
# spire.xls c# excel Worksheets
## Find All Formulas in Worksheet
//Find cells containing the specific formula "=SUM(A1:A5)"
Spire.Xls.CellRange[] formulaRanges = sheet.FindAll("=SUM(A1:A5)", Spire.Xls.FindType.Formula, Spire.Xls.ExcelFindOptions.None);
if (formulaRanges != null)
{
    foreach (Spire.Xls.CellRange range in formulaRanges)
    {
        Console.WriteLine("Formula found at: " + range.RangeAddress);
    }
}

# spire.xls c# excel Worksheets
## Get Cell Data Type
//Get type of cell B2
Spire.Xls.Core.Spreadsheet.Xlssheet.TRangeValueType cellType = sheet.GetCellType(2, 2, false); // (row, col, evaluateFormula)
Console.WriteLine("Cell B2 type: " + cellType.ToString());

# spire.xls c# excel Worksheets
## Get Active Selection Range in Worksheet
var activeSelection = sheet.ActiveSelectionRange;
if (activeSelection != null)
{
    foreach (Spire.Xls.CellRange selectedRange in activeSelection)
    {
        Console.WriteLine("Active selection: " + selectedRange.RangeAddress);
    }
}

# spire.xls c# excel Worksheets
## Ungroup Rows in Worksheet
//Assume rows 10 to 12 were previously grouped
sheet.UngroupByRows(10, 12);

# spire.xls c# excel Worksheets
## Set Column Width
//Set width of column B (index 2) to 20 characters
sheet.SetColumnWidth(2, 20.0);

# spire.xls c# excel Worksheets
## AutoFit Specific Columns in Worksheet
//AutoFit columns B to D (indices 2 to 4)
sheet.AutoFitColumn(2, 2, 4); // (firstColumn, lastColumn)
//To AutoFit a single column:
sheet.AutoFitColumn(1); // AutoFit column A

# spire.xls c# excel Worksheets
## AutoFit Specific Rows in Worksheet
//AutoFit rows 2 to 5, considering merged cells
sheet.AutoFitRow(2, 1, 5, true); // (rowIndex, firstColumn, lastColumn, bRaiseEvents)
sheet.AutoFitRow(2); // AutoFit row 2

# spire.xls c# excel Worksheets
## Check if Row is AutoFit
bool isRow2AutoFit = sheet.GetRowIsAutoFit(2);
Console.WriteLine("Row 2 is AutoFit: " + isRow2AutoFit);

# spire.xls c# excel Worksheets
## Check if Column is AutoFit
bool isColumnBAutoFit = sheet.GetColumnIsAutoFit(2); // Column B (index 2)
Console.WriteLine("Column B is AutoFit: " + isColumnBAutoFit);

# spire.xls c# excel Worksheets
## Check if Row is Hidden
int rowIndexToCheck = 3; // Check row 3
bool isRowHidden = sheet.GetRowIsHide(rowIndexToCheck);
Console.WriteLine("Row " + rowIndexToCheck + " is hidden: " + isRowHidden);

# spire.xls c# excel Worksheets
## Check if Column is Hidden
int columnIndexToCheck = 2; // Check column B (index 2)
bool isColumnHidden = sheet.GetColumnIsHide(columnIndexToCheck);
Console.WriteLine("Column " + columnIndexToCheck + " is hidden: " + isColumnHidden);

# spire.xls c# excel Worksheets
## Delete Row from Worksheet
//Delete row 5
sheet.DeleteRow(5);
//Delete 3 rows starting from row 2
sheet.DeleteRow(2, 3);

# spire.xls c# excel Worksheets
## Delete Column from Worksheet
//Delete column C (index 3)
sheet.DeleteColumn(3);
//Delete 2 columns starting from column B (index 2)
sheet.DeleteColumn(2, 2);

# spire.xls c# excel Worksheets
## Get Total Row Count in Worksheet
int totalRows = sheet.Rows.Length;
Console.WriteLine("Total rows in worksheet: " + totalRows);

# spire.xls c# excel Worksheets
## Get Total Column Count in Worksheet
int totalColumns = sheet.Columns.Length;
Console.WriteLine("Total columns in worksheet: " + totalColumns);

# spire.xls c# excel Worksheets
## Group Columns in Worksheet
//Group columns B to D (indices 2 to 4), do not collapse them by default
sheet.GroupByColumns(2, 4, false);

# spire.xls c# excel Worksheets
## Set Row and Column Headers Visibility
sheet.RowColumnHeadersVisible = false; // Set to true to show them

# spire.xls c# excel Worksheets
## Hide Column in Worksheet
//Hide column B (index 2)
sheet.HideColumn(2);

# spire.xls c# excel Worksheets
## Hide Row in Worksheet
//Hide row 4
sheet.HideRow(4);

# spire.xls c# excel Worksheets
## Insert Row into Worksheet
//Insert a new row at index 2 (becomes the new row 2)
sheet.InsertRow(2);
//Insert 3 new rows starting at index 5
sheet.InsertRow(5, 3);

# spire.xls c# excel Worksheets
## Insert Column into Worksheet
//Insert a new column at index 2 (becomes the new column B)
sheet.InsertColumn(2);
//Insert 2 new columns starting at index 4
sheet.InsertColumn(4, 2);

# spire.xls c# excel Worksheets
## Set Column Width in Pixels
//Set width of column C (index 3) to 100 pixels
sheet.SetColumnWidthInPixels(3, 100);

# spire.xls c# excel Worksheets
## Set Default Column Width
sheet.DefaultColumnWidth = 15; // Default width in characters

# spire.xls c# excel Worksheets
## Set Default Row Style
Spire.Xls.CellStyle defaultRowStyle = sheet.Workbook.Styles.Add("DefaultRowStyle");
defaultRowStyle.Font.Color = System.Drawing.Color.DarkGreen;
defaultRowStyle.Interior.Color = System.Drawing.Color.FromArgb(230, 255, 230); // Light green background
//Set default style for row 1 (first row)
sheet.SetDefaultRowStyle(1, defaultRowStyle);

# spire.xls c# excel Worksheets
## Set Default Column Style
Spire.Xls.CellStyle defaultColStyle = sheet.Workbook.Styles.Add("DefaultColumnStyle");
defaultColStyle.Font.IsBold = true;
defaultColStyle.Font.Color = System.Drawing.Color.DarkBlue;
//Set default style for column A (first column, index 1)
sheet.SetDefaultColumnStyle(1, defaultColStyle);

# spire.xls c# excel Worksheets
## Set Default Row Height
sheet.DefaultRowHeight = 20; // Default height in points

# spire.xls c# excel Worksheets
## Set Row Height
//Set height of row 4 to 30 points
sheet.SetRowHeight(4, 30.0);

# spire.xls c# excel Worksheets
## Set PageSetup SummaryColumnRight Property
//This affects outlining when summary columns are present
sheet.PageSetup.IsSummaryColumnRight = true; // True to display summary columns to the right of detail

# spire.xls c# excel Worksheets
## Show Hidden Row
//Assume row 15 was previously hidden
sheet.ShowRow(15);

# spire.xls c# excel Worksheets
## Show Hidden Column
//Assume column D (index 4) was previously hidden
sheet.ShowColumn(4);

# spire.xls c# excel Worksheets
## Add Picture to Worksheet
string imagePath = @"path\to\your\image.png"; // Replace with actual path
if (System.IO.File.Exists(imagePath))
{
    Spire.Xls.ExcelPicture picture = sheet.Pictures.Add(3, 3, imagePath); // (topRow, leftColumn, filePath)
    picture.Width = 100; // Optional: set picture properties
    picture.Height = 80;
}

# spire.xls c# excel Worksheets
## Remove Picture from Worksheet by Index
if (sheet.Pictures.Count > 0)
{
    sheet.Pictures.RemoveAt(0); // Removes the first picture
}

# spire.xls c# excel Worksheets
## Get Embedded Cell Images from Worksheet
Spire.Xls.ExcelPicture[] cellImages = sheet.CellImages;
if (cellImages != null)
{
    foreach (Spire.Xls.ExcelPicture pic in cellImages)
    {
        Console.WriteLine("Found embedded cell image: " + pic.FileName);
        pic.SaveToImage("embedded_" + pic.FileName);
    }
}

# spire.xls c# excel Worksheets
## Set Worksheet Background Image
string backgroundImagePath = @"path\to\your\background.jpg"; // Replace with actual path
if (System.IO.File.Exists(backgroundImagePath))
{
    System.IO.FileStream stream = new System.IO.FileStream(backgroundImagePath,FileMode.OpenOrCreate);
    SkiaSharp.SKBitmap backgroundImage = SkiaSharp.SKBitmap.Decode(stream);
    sheet.PageSetup.BackgoundImage = backgroundImage;
}

# spire.xls c# excel Worksheets
## Add Chart to Worksheet
//Assume data for chart in A1:B5
//sheet.Range["A1"].Value = "Category"; sheet.Range["B1"].Value = "Value"; ...
Chart chart = sheet.Charts.Add(Spire.Xls.ExcelChartType.ColumnClustered);
chart.DataRange = sheet.Range["A1:B5"]; // Set data range for the chart
chart.SeriesDataFromRange = false;     // Data is in columns for categories, series in rows
chart.ChartTitle = "My Chart";
// Position the chart
chart.LeftColumn = 4;
chart.TopRow = 2;
chart.RightColumn = 12;
chart.BottomRow = 15;

# spire.xls c# excel Worksheets
## Get Comment from Worksheet by Index
if (sheet.Comments.Count > 0)
{
    Spire.Xls.ExcelComment comment = sheet.Comments[0] ; // Gets the first comment
    Console.WriteLine("First comment text: " + comment.Text);
}

# spire.xls c# excel Worksheets
## Get All Comments from Worksheet
Spire.Xls.Collections.CommentsCollection comments = sheet.Comments;
foreach (var comment in comments)
{
    Console.WriteLine("Comment by " + comment.Author + " at " + comment.LinkedCell.RangeAddress + ": " + comment.Text);
}


# spire.xls c# excel Worksheets
## Save Worksheet to PDF
string outputPdfPath = "WorksheetOutput.pdf";
sheet.SaveToPdf(outputPdfPath);
Console.WriteLine("Worksheet saved to PDF: " + outputPdfPath);

# spire.xls c# excel Worksheets
## Set PageSetup FitToPagesWide Property
//Fit content to 1 page wide when printing or converting to PDF
sheet.PageSetup.FitToPagesWide = 1;
//To disable fitting, set to 0
sheet.PageSetup.FitToPagesTall = 0; // Optionally set height fitting

# spire.xls c# excel Worksheets
## Save Worksheet to Image
//Save the entire used range of the worksheet to an image stream
Stream sheetImage = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn);
// sheetImage.Save("worksheet_image.png", System.Drawing.Imaging.ImageFormat.Png);
Console.WriteLine("Worksheet converted to image.");

# spire.xls c# excel Worksheets
## Save Worksheet to Text File
string outputTextPath = "WorksheetOutput.txt";
//Save worksheet to a text file, using tab as delimiter
sheet.SaveToFile(outputTextPath, "\t", System.Text.Encoding.UTF8);
Console.WriteLine("Worksheet saved to text file: " + outputTextPath);

# spire.xls c# excel Worksheets
## Save Worksheet to HTML with Options
Spire.Xls.Core.Spreadsheet.HTMLOptions htmlOptions = new Spire.Xls.Core.Spreadsheet.HTMLOptions();
htmlOptions.ImageEmbedded = true; // Embed images in the HTML
// Set the HTMLOptions to create a standalone HTML file.
htmlOptions.IsStandAloneHtmlFile = true; 
string outputHtmlPath = "WorksheetOutput.html";
sheet.SaveToHtml(outputHtmlPath, htmlOptions);
Console.WriteLine("Worksheet saved to HTML: " + outputHtmlPath);

# spire.xls c# excel Worksheets
## Save Worksheet to HTML Stream with Options
var htmlOptions = new Spire.Xls.Core.Spreadsheet.HTMLOptions();
htmlOptions.ImageEmbedded = false; // Link images instead of embedding
using (System.IO.FileStream htmlStream = new System.IO.FileStream("WorksheetStreamOutput.html", System.IO.FileMode.Create))
{
    sheet.SaveToHtml(htmlStream, htmlOptions);
}
Console.WriteLine("Worksheet saved to HTML stream.");

# spire.xls c# excel Worksheets
## Add Conditional Formats to Worksheet
Spire.Xls.Core.Spreadsheet.Collections.XlsConditionalFormats conditionalFormats = sheet.ConditionalFormats.Add();
//Define the range to apply conditional formatting
conditionalFormats.AddRange(sheet.Range["A1:A10"]);
//Add a condition (e.g., highlight cells with value > 50)
Spire.Xls.Core.IConditionalFormat condition = conditionalFormats.AddCondition();
condition.FormatType = Spire.Xls.ConditionalFormatType.CellValue;
condition.Operator = Spire.Xls.ComparisonOperatorType.Greater;
condition.FirstFormula = "50";
condition.BackColor = System.Drawing.Color.LightYellow;

# spire.xls c# excel Worksheets
## Activate Worksheet
sheet.Activate();
Console.WriteLine("Worksheet '" + sheet.Name + "' activated.");

# spire.xls c# excel Worksheets
## Set Active Cell in Worksheet
//Set cell C5 as the active cell
sheet.SetActiveCell(sheet.Range["C5"]);
Console.WriteLine("Cell C5 is now active in sheet '" + sheet.Name + "'.");

# spire.xls c# excel Worksheets
## Set First Visible Column in Worksheet
//Make column C (index 3) the first visible column
sheet.FirstVisibleColumn = 3;

# spire.xls c# excel Worksheets
## Set First Visible Row in Worksheet
//Make row 5 the first visible row
sheet.FirstVisibleRow = 5;

# spire.xls c# excel Worksheets
## Set Different First Page Header/Footer Behavior
//Enable different header/footer for the first page
sheet.PageSetup.DifferentFirst = 1; // Use 0 to disable
//Then set FirstHeaderString, FirstFooterString etc.
sheet.PageSetup.FirstHeaderString = "First Header";

# spire.xls c# excel Worksheets
## Set First Page Left Header Image
string imagePath = @"path\to\your\header_image.png"; // Replace with actual path
if (System.IO.File.Exists(imagePath))
{
    System.IO.FileStream stream = new System.IO.FileStream(imagePath, FileMode.OpenOrCreate);
    SkiaSharp.SKBitmap headerImage = SkiaSharp.SKBitmap.Decode(stream);
    sheet.PageSetup.DifferentFirst = 1; // Ensure different first page is enabled
    sheet.PageSetup.FirstLeftHeaderImage = headerImage;
    sheet.PageSetup.LeftHeader = "&G"; // Command to display image
}

# spire.xls c# excel Worksheets
## Set Worksheet View Mode
sheet.ViewMode = Spire.Xls.ViewMode.Layout; // Other modes: Normal, Preview
Console.WriteLine("Worksheet view mode set to Layout.");
# spire.xls c# excel Worksheets
## Set PageSetup Left Header Image
string imagePath = @"path\to\your\page_header_image.png"; // Replace with actual path
if (System.IO.File.Exists(imagePath))
{
    System.IO.FileStream stream = new System.IO.FileStream(imagePath, FileMode.OpenOrCreate);
    SkiaSharp.SKBitmap pageHeaderImage = SkiaSharp.SKBitmap.Decode(stream);
    sheet.PageSetup.LeftHeaderImage = pageHeaderImage;
    sheet.PageSetup.LeftHeader = "&G"; // Command to display image
}

# spire.xls c# excel Worksheets
## Set PageSetup Left Header Text
sheet.PageSetup.LeftHeader = "Company Confidential";
//For font styling: "&\"Arial,Bold\"&12Company Confidential"

# spire.xls c# excel Worksheets
## Set Different Odd/Even Page Header/Footer Behavior
//Enable different headers/footers for odd and even pages
sheet.PageSetup.DifferentOddEven = 1; // Use 0 to disable
//Then set OddHeaderString, EvenFooterString etc.
sheet.PageSetup.OddHeaderString = "Odd Page Header";
sheet.PageSetup.EvenHeaderString = "Even Page Header";

# spire.xls c# excel Worksheets
## Set Odd Page Header String
sheet.PageSetup.DifferentOddEven = 1; // Ensure different odd/even is enabled
sheet.PageSetup.OddHeaderString = "&\"Verdana,Bold Italic\"&10Odd Page - Title";
//Similarly for EvenHeaderString, OddFooterString, EvenFooterString

# spire.xls c# excel Worksheets
## Set First Page Header String
sheet.PageSetup.DifferentFirst = 1; // Ensure different first page is enabled
sheet.PageSetup.FirstHeaderString = "&C&\"Times New Roman,Bold\"&14First Page Main Title";
//Similarly for FirstFooterString

# spire.xls c# excel Worksheets
## Set PageSetup Header/Footer Picture Crop
//Assume an image is already set for LeftHeader
//sheet.PageSetup.LeftHeaderImage = someImage;
//sheet.PageSetup.LeftHeader = "&G";
//Crop 20% from top, 30% from bottom, 30% from left, 20% from right for the left header picture
sheet.PageSetup.LeftHeaderPictureCropTop = 0.2f;
sheet.PageSetup.LeftHeaderPictureCropBottom = 0.3f;
sheet.PageSetup.LeftHeaderPictureCropLeft = 0.3f;
sheet.PageSetup.LeftHeaderPictureCropRight = 0.2f;
//Similar properties exist for CenterHeaderPictureCrop, RightHeaderPictureCrop, and Footer pictures

# spire.xls c# excel Worksheets
## Add Hyperlink to Cell Range
Spire.Xls.CellRange cellWithLink = sheet.Range["A1"];
cellWithLink.Text = "Visit Spire.XLS";
Spire.Xls.HyperLink hyperlink = sheet.HyperLinks.Add(cellWithLink);
hyperlink.Type = Spire.Xls.HyperLinkType.Url;
hyperlink.Address = "https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html";

# spire.xls c# excel Worksheets
## Iterate Through Hyperlinks in Worksheet
if (sheet.HyperLinks.Count > 0)
{
    foreach (Spire.Xls.HyperLink link in sheet.HyperLinks)
    {
        Console.WriteLine("Link Text: " + link.TextToDisplay + ", Address: " + link.Address + ", Type: " + link.Type);
    }
}

# spire.xls c# excel Worksheets
## Remove Hyperlink from Worksheet by Index
if (sheet.HyperLinks.Count > 0)
{
    sheet.HyperLinks.RemoveAt(0); // Removes the first hyperlink
    Console.WriteLine("First hyperlink removed.");
}

# spire.xls c# excel Worksheets
## Check if Worksheet Has OLE Objects
bool hasOle = sheet.HasOleObjects;
Console.WriteLine("Worksheet has OLE objects: " + hasOle);

# spire.xls c# excel Worksheets
## Get OLE Object from Worksheet by Index
if (sheet.HasOleObjects && sheet.OleObjects.Count > 0)
{
    Spire.Xls.Core.IOleObject oleObject = sheet.OleObjects[0] ; // Gets the first OLE object
    Console.WriteLine("First OLE object type: " + oleObject.ObjectType);
}

# spire.xls c# excel Worksheets
## Add OLE Object to Worksheet
string oleFilePath = @"path\to\your\document.docx"; // Example: Word document
string iconImagePath = @"path\to\your\icon.png";   // Example: Icon for OLE object
if (System.IO.File.Exists(oleFilePath) && System.IO.File.Exists(iconImagePath))
{
    System.IO.FileStream stream = new System.IO.FileStream(oleFilePath, FileMode.OpenOrCreate);
    SkiaSharp.SKBitmap displayIcon = SkiaSharp.SKBitmap.Decode(stream);
    Spire.Xls.Core.IOleObject oleObject = sheet.OleObjects.Add(oleFilePath, displayIcon, Spire.Xls.OleLinkType.Embed);
    oleObject.Location = sheet.Range["C5"]; // Position the OLE object
    oleObject.ObjectType = Spire.Xls.OleObjectType.WordDocument; // Set the type if known
}

# spire.xls c# excel Worksheets
## Set PageSetup Paper Size
sheet.PageSetup.PaperSize = Spire.Xls.PaperSizeType.PaperA4;

# spire.xls c# excel Worksheets
## Get PageSetup Page Width and Height
double pageWidth = sheet.PageSetup.PageWidth; // In points
double pageHeight = sheet.PageSetup.PageHeight; // In points
Console.WriteLine("Page dimensions: " + pageWidth + "pt x " + pageHeight + "pt");

# spire.xls c# excel Worksheets
## Set PageSetup Page Order
//Order for printing multiple pages: OverThenDown or DownThenOver
sheet.PageSetup.Order = Spire.Xls.OrderType.OverThenDown;

# spire.xls c# excel Worksheets
## Set PageSetup First Page Number
//Set the starting page number for printing this sheet
sheet.PageSetup.FirstPageNumber = 5;

# spire.xls c# excel Worksheets
## Set PageSetup Header/Footer Margin (Inch)
sheet.PageSetup.HeaderMarginInch = 0.5; // 0.5 inch header margin
sheet.PageSetup.FooterMarginInch = 0.5; // 0.5 inch footer margin

# spire.xls c# excel Worksheets
## Set PageSetup Page Margins (Inch)
sheet.PageSetup.TopMargin = 0.75;    // 0.75 inch
sheet.PageSetup.BottomMargin = 0.75; // 0.75 inch
sheet.PageSetup.LeftMargin = 0.7;    // 0.7 inch
sheet.PageSetup.RightMargin = 0.7;   // 0.7 inch

# spire.xls c# excel Worksheets
## Set PageSetup Print Gridlines
sheet.PageSetup.IsPrintGridlines = true; // Set to false to hide gridlines on print

# spire.xls c# excel Worksheets
## Set PageSetup Print Comments Type
sheet.PageSetup.PrintComments = Spire.Xls.PrintCommentType.InPlace; // Options: DisplayAtEnd, InPlace, PrintNoComments

# spire.xls c# excel Worksheets
## Set PageSetup Print Errors Type
sheet.PageSetup.PrintErrors = Spire.Xls.PrintErrorsType.NA; // Options: PrintErrorsBlank, Dash, Displayed, NA

# spire.xls c# excel Worksheets
## Set PageSetup Draft Quality Printing
sheet.PageSetup.Draft = true; // Set to false for normal quality

# spire.xls c# excel Worksheets
## Set PageSetup Print Headings
//Print row and column headings (e.g., A, B, C and 1, 2, 3)
sheet.PageSetup.IsPrintHeadings = true; // Set to false to hide them

# spire.xls c# excel Worksheets
## Set PageSetup Black and White Printing
sheet.PageSetup.BlackAndWhite = true; // Set to false for color printing

# spire.xls c# excel Worksheets
## Set PageSetup Page Orientation
sheet.PageSetup.Orientation = Spire.Xls.PageOrientationType.Landscape; // Or PageOrientationType.Portrait

# spire.xls c# excel Worksheets
## Set PageSetup Print Area
//Set the print area to cells A1 through D20
sheet.PageSetup.PrintArea = "A1:D20";
//To clear print area:
sheet.PageSetup.PrintArea = null;

# spire.xls c# excel Worksheets
## Set PageSetup Print Quality
//Set print quality in dots per inch (DPI). Common values might be 72, 150, 300, 600.
sheet.PageSetup.PrintQuality = 300; // Example: 300 DPI

# spire.xls c# excel Worksheets
## Set PageSetup Print Title Columns
//Repeat columns A and B on every printed page
sheet.PageSetup.PrintTitleColumns = "$A:$B";

# spire.xls c# excel Worksheets
## Set PageSetup Print Title Rows
//Repeat rows 1 and 2 on every printed page
sheet.PageSetup.PrintTitleRows = "$1:$2";

# spire.xls c# excel Worksheets
## Set PageSetup Fit To Pages Tall
//Fit content to 2 pages tall when printing
sheet.PageSetup.FitToPagesTall = 2;
//To disable fitting, set to 0
sheet.PageSetup.FitToPagesWide = 0; // Optionally set width fitting

# spire.xls c# excel Worksheets
## Set PageSetup Center Horizontally
//Center the worksheet content horizontally on the printed page
sheet.PageSetup.CenterHorizontally = true;

# spire.xls c# excel Worksheets
## Set PageSetup Center Vertically
//Center the worksheet content vertically on the printed page
sheet.PageSetup.CenterVertically = true;

# spire.xls c# excel Worksheets
## Get PivotTable from Worksheet by Name/Index
//Get by index (first PivotTable)
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTableByIndex = sheet.PivotTables[0] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;
if (pivotTableByIndex != null)
{
    Console.WriteLine("PivotTable by index: " + pivotTableByIndex.Name);
}
//Get by name (assuming a PivotTable named "SalesPivot")
Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable pivotTableByName = sheet.PivotTables["SalesPivot"] as Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable;

# spire.xls c# excel Worksheets
## Add PivotTable to Worksheet//Assume 'Worksheet sheet' for PivotTable destination exists
//Assume 'Worksheet dataSourceSheet' for data source exists
//Assume 'Spire.Xls.PivotCache cache' is created from dataSourceSheet, e.g.,
Spire.Xls.CellRange dataRange = dataSourceSheet.Range["A1:D100"];
Spire.Xls.PivotCache cache = sheet.Workbook.PivotCaches.Add(dataRange);
Spire.Xls.PivotCache pivotCache = sheet.Workbook.PivotCaches.Add(dataSourceSheet.Range["A1:D10"]); // Example
Spire.Xls.PivotTable pivotTable = sheet.PivotTables.Add("NewPivotTable", sheet.Range["G1"], pivotCache);
//Configure pivotTable fields (row, column, data fields)
pivotTable.PivotFields["Category"].Axis = Spire.Xls.AxisTypes.Row;
pivotTable.DataFields.Add(pivotTable.PivotFields["Sales"], "Sum of Sales", Spire.Xls.SubtotalTypes.Sum);
pivotTable.CalculateData();

# spire.xls c# excel Worksheets
## Set PageSetup Custom Paper Size Name
//This name should correspond to a custom paper size defined in the printer settings.
sheet.PageSetup.CustomPaperSizeName = "MyCustomPaper";

# spire.xls c# excel Worksheets
## Set PageSetup Custom Paper Size Dimensions
//Set custom paper size: width 200 points, height 300 points
sheet.PageSetup.SetCustomPaperSize(200f, 300f);

# spire.xls c# excel Worksheets
## Protect Worksheet with Password and Protection Type
//Protect sheet, allowing only selection of unlocked cells
sheet.Protect("myPassword123", Spire.Xls.SheetProtectionType.UnLockedCells);
// To protect all aspects: Spire.Xls.SheetProtectionType.All

# spire.xls c# excel Worksheets
## Add Allow Edit Range to Protected Worksheet
//First, protect the sheet
sheet.Protect("password", Spire.Xls.SheetProtectionType.All);
//Then, add a range that users can edit even when the sheet is protected
sheet.AddAllowEditRange("EditableRegion1", sheet.Range["B2:D5"]);

# spire.xls c# excel Worksheets
## Unprotect Worksheet with Password
sheet.Unprotect("myPassword123");
Console.WriteLine("Worksheet unprotected with password.");

# spire.xls c# excel Worksheets
## Unprotect Worksheet (No Password)
sheet.Unprotect();
Console.WriteLine("Worksheet unprotected (no password).");

# spire.xls c# excel Worksheets
## Add Line Shape to Worksheet (TypedLines)
Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape line = sheet.TypedLines.AddLine() as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
line.TopRow = 2;       // Start row (0-indexed)
line.LeftColumn = 2;    // Start column (0-indexed)
line.BottomRow = 5;         // End row
line.RightColumn = 5;      // End column
line.LineShapeType = LineShapeType.ElbowLine;
line.Width = 30;
line.Height = 50;
line.Color = System.Drawing.Color.Blue;
line.EndArrowHeadStyle = Spire.Xls.ShapeArrowStyleType.LineArrow;

# spire.xls c# excel Worksheets
## Add Line Shape to Worksheet (Lines Collection)
//Add a straight line from (row 10, col 2, height 200, width 1 - effectively a vertical line segment)
Spire.Xls.Core.ILineShape line = sheet.Lines.AddLine(10, 2, 200, 1, Spire.Xls.LineShapeType.Line);
line.DashStyle = Spire.Xls.ShapeDashLineStyleType.Solid;
line.Color = System.Drawing.Color.Red;
line.Weight = 2f;

# spire.xls c# excel Worksheets
## Add Oval Shape to Worksheet
Spire.Xls.Core.IOvalShape oval = sheet.OvalShapes.AddOval(5, 3, 100, 150); // (topRow, leftColumn, height, width) in default units
oval.Fill.ForeColor = System.Drawing.Color.LightSkyBlue;
oval.Fill.FillType = Spire.Xls.ShapeFillType.SolidColor;
oval.Line.ForeColor = System.Drawing.Color.DarkBlue;

# spire.xls c# excel Worksheets
## Add Rectangle Shape to Worksheet
Spire.Xls.Core.IRectangleShape rect = sheet.RectangleShapes.AddRectangle(6, 2, 80, 200, RectangleShapeType.Rect); // (topRow, leftColumn, height, width, type)
rect.Fill.ForeColor = System.Drawing.Color.PaleGreen;
rect.Fill.FillType = Spire.Xls.ShapeFillType.SolidColor;

# spire.xls c# excel Worksheets
## Add Spinner Shape to Worksheet
//Link to cell C12
//sheet.Range["C12"].Value2 = 0;
Spire.Xls.Core.ISpinnerShape spinner = sheet.SpinnerShapes.AddSpinner(12, 4, 20, 20); // (topRow, leftColumn, height, width)
spinner.LinkedCell = sheet.Range["C12"];
spinner.Min = 0;
spinner.Max = 100;
spinner.IncrementalChange = 5;

# spire.xls c# excel Worksheets
## Save All Shapes in Worksheet to Images
Spire.Xls.SaveShapeTypeOption sstOptions = new Spire.Xls.SaveShapeTypeOption();
sstOptions.SaveAll = true; // Save all types of shapes
//sstOptions.SaveTextBox = true; // To save only textboxes
List<SkiaSharp.SKBitmap> shapeImages = sheet.SaveShapesToImage(sstOptions);
int i = 0;
foreach (SkiaSharp.SKBitmap img in shapeImages)
{
}
Console.WriteLine(shapeImages.Count + " shapes saved as images.");

# spire.xls c# excel Worksheets
## Copy Line Shape to Another Worksheet
//Assume 'Worksheet destinationSheet' exists
if (sourceSheet.TypedLines.Count > 0)
{
    Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape lineToCopy = sourceSheet.TypedLines[0] as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
    destinationSheet.TypedLines.AddCopy(lineToCopy);
    Console.WriteLine("Line shape copied to destination sheet.");
}
// Similar AddCopy methods exist for TypedTextBoxes, TypedCheckBoxes, etc.

# spire.xls c# excel Worksheets
## Remove Preset Geometry Shape from Worksheet by Index
if (sheet.PrstGeomShapes.Count > 0)
{
    sheet.PrstGeomShapes.RemoveAt(0); // Removes the first preset geometry shape
    Console.WriteLine("First preset geometry shape removed.");
}

# spire.xls c# excel Worksheets
## Get Preset Geometry Shape from Worksheet by Index
if (sheet.PrstGeomShapes.Count > 0)
{
    Spire.Xls.Core.IPrstGeomShape shape = sheet.PrstGeomShapes[0] ; // Gets the first preset geometry shape
    Console.WriteLine("First preset geometry shape type: " + shape.PrstShapeType);
}

# spire.xls c# excel Worksheets
## Get Collection of Preset Geometry Shapes
Spire.Xls.Core.Spreadsheet.Collections.PrstGeomShapeCollection shapes = sheet.PrstGeomShapes;
Console.WriteLine("Number of preset geometry shapes: " + shapes.Count);
foreach (Spire.Xls.Core.IPrstGeomShape shape in shapes)
{
    // Process each shape
}

# spire.xls c# excel Worksheets
## Add Preset Geometry Shape to Worksheet
Spire.Xls.Core.IPrstGeomShape heartShape = sheet.PrstGeomShapes.AddPrstGeomShape(
    5, 2, 100, 100, Spire.Xls.PrstGeomShapeType.Heart); // (topRow, leftColumn, height, width, shapeType)
heartShape.Fill.ForeColor = System.Drawing.Color.Red;
heartShape.Fill.FillType = Spire.Xls.ShapeFillType.SolidColor;

# spire.xls c# excel Worksheets
## Get GroupShapeCollection from Worksheet
Spire.Xls.Core.MergeSpreadsheet.Collections.GroupShapeCollection groupShapes = sheet.GroupShapeCollection;
// To group shapes:
IPrstGeomShape shape1 = sheet.PrstGeomShapes.AddPrstGeomShape(1, 3, 50, 50, PrstGeomShapeType.RoundRect);
IPrstGeomShape shape2 = sheet.PrstGeomShapes.AddPrstGeomShape(5, 3, 50, 50, PrstGeomShapeType.Triangle);
groupShapes.Group(new Spire.Xls.Core.IShape[] { shape1, shape2 });
Console.WriteLine("Number of grouped shapes/groups: " + groupShapes.Count);

# spire.xls c# excel Worksheets
## Save and Get Shapes to Images from Worksheet
Spire.Xls.SaveShapeTypeOption sstOptions = new Spire.Xls.SaveShapeTypeOption();
sstOptions.SaveAll = true;
var shapeImageMap = sheet.SaveAndGetShapesToImage(sstOptions);
foreach (System.Collections.Generic.KeyValuePair<Spire.Xls.Core.IShape, SkiaSharp.SKBitmap> entry in shapeImageMap)
{
    Spire.Xls.Core.IShape shape = entry.Key;
    SkiaSharp.SKBitmap image = entry.Value;
    Console.WriteLine("Shape: " + shape.Name + ", Image Size: " + image.Width + "x" + image.Height);
    // image.Save(shape.Name + ".png");
}

# spire.xls c# excel Worksheets
## Get TextBox from Worksheet by Index
if (sheet.TextBoxes.Count > 0)
{
    Spire.Xls.Core.ITextBoxShape textBox = sheet.TextBoxes[0] ; // Gets the first textbox
    Console.WriteLine("First textbox text: " + textBox.Text);
}

# spire.xls c# excel Worksheets
## Get TextBox from Worksheet by Name
Spire.Xls.Core.ITextBoxShape namedTextBox = sheet.TextBoxes["InfoBox"];
if (namedTextBox != null)
{
    Console.WriteLine("Text in 'InfoBox': " + namedTextBox.Text);
}

# spire.xls c# excel Worksheets
## Add Horizontal Page Break to Worksheet
//Add a horizontal page break above row 10
Spire.Xls.HPageBreak pageBreak = sheet.HPageBreaks.Add(sheet.Range["A10"]);
Console.WriteLine("Horizontal page break added at row " + pageBreak.Row);

# spire.xls c# excel Worksheets
## Add Vertical Page Break to Worksheet
//Add a vertical page break to the left of column E
Spire.Xls.VPageBreak vPageBreak = sheet.VPageBreaks.Add(sheet.Range["E1"]);
Console.WriteLine("Vertical page break added at column " + vPageBreak.Column);
# spire.xls c# excel Worksheets
## Copy Worksheet Content from Another Worksheet (CopyFrom)
//Assume 'Worksheet sourceSheetToCopy' exists and contains data
destinationSheet.CopyFrom(sourceSheetToCopy);
Console.WriteLine("Content from '" + sourceSheetToCopy.Name + "' copied to '" + destinationSheet.Name + "'.");

# spire.xls c# excel Worksheets
## Check Worksheet Visibility
Spire.Xls.WorksheetVisibility visibility = sheet.Visibility;
Console.WriteLine("Worksheet '" + sheet.Name + "' visibility: " + visibility.ToString());
// To set: sheet.Visibility = Spire.Xls.WorksheetVisibility.Hidden;

# spire.xls c# excel Worksheets
## Check if Worksheet is Empty
bool isEmpty = sheet.IsEmpty;
Console.WriteLine("Worksheet '" + sheet.Name + "' is empty: " + isEmpty);
# spire.xls c# excel Worksheets
## Freeze Panes in Worksheet
//Freeze the first row and first column (panes split at cell B2)
sheet.FreezePanes(2, 2); // (rowIndex, columnIndex) - 1-based index for user view, so 2,2 means cell B2
Console.WriteLine("Panes frozen at B2.");

# spire.xls c# excel Worksheets
## Get Custom Properties of Worksheet
ICustomPropertiesCollection customProps = sheet.CustomProperties;
if (customProps.Count > 0)
{
    foreach (XlsCustomProperty prop in customProps)
    {
        Console.WriteLine("Custom Property: " + prop.Name + " = " + prop.Value);
    }
}
else
{
    Console.WriteLine("Worksheet has no custom properties.");
}

// To add: sheet.CustomProperties.Add("MyProp", "MyValue");
# spire.xls c# excel Worksheets
## Get Freeze Pane Row and Column Index
int frozenRow, frozenCol;
sheet.GetFreezePanes(out frozenRow, out frozenCol);
if (frozenRow > 0 || frozenCol > 0)
{
    Console.WriteLine("Panes frozen at row " + frozenRow + ", column " + frozenCol + " (0-indexed if not frozen, 1-indexed if frozen).");
}
else
{
    Console.WriteLine("No panes are frozen.");
}

# spire.xls c# excel Worksheets
## Set Worksheet Visibility
sheet.Visibility = Spire.Xls.WorksheetVisibility.Hidden; // Options: Visible, Hidden, VeryHidden
Console.WriteLine("Worksheet '" + sheet.Name + "' is now hidden.");

# spire.xls c# excel Worksheets
## Set Display Zeros in Worksheet
//Set to false to hide zero values in cells
sheet.IsDisplayZeros = false;
Console.WriteLine("Display of zero values is now turned off for sheet '" + sheet.Name + "'.");

# spire.xls c# excel Worksheets
## Move Worksheet to New Position
//Move the sheet to be the third sheet in the workbook (index 2)
sheet.MoveWorksheet(2); // 0-based index for new position
Console.WriteLine("Worksheet '" + sheet.Name + "' moved to position 2.");

# spire.xls c# excel Worksheets
## Set Zoom Scale for Page Break Preview
//Set the zoom scale for the Page Break Preview mode to 75%
sheet.ZoomScalePageBreakView = 75;
//Ensure view mode is PageBreakPreview to see effect:
//sheet.ViewMode = Spire.Xls.ViewMode.Preview;

# spire.xls c# excel Worksheets
## Clear All Vertical Page Breaks
sheet.VPageBreaks.Clear();
Console.WriteLine("All vertical page breaks cleared.");

# spire.xls c# excel Worksheets
## Remove Horizontal Page Break by Index
if (sheet.HPageBreaks.Count > 0)
{
    sheet.HPageBreaks.RemoveAt(0); // Removes the first horizontal page break
    Console.WriteLine("First horizontal page break removed.");
}

# spire.xls c# excel Worksheets
## Set Worksheet Tab Color
sheet.TabColor = System.Drawing.Color.DodgerBlue;

# spire.xls c# excel Worksheets
## Set Grid Lines Visibility
sheet.GridLinesVisible = false; // Set to true to show grid lines
Console.WriteLine("Grid lines visibility set to false for sheet '" + sheet.Name + "'.");

# spire.xls c# excel Worksheets
## Set Vertical Split Position
//Set vertical split at 2000 twips (twentieths of a point) from the left
sheet.VerticalSplit = 2000;
//Also requires setting FirstVisibleColumn if you want a specific column to be active in the left pane.
sheet.FirstVisibleColumn = 3;

# spire.xls c# excel Worksheets
## Set Horizontal Split Position
//Set horizontal split at 3000 twips (twentieths of a point) from the top
sheet.HorizontalSplit = 3000;
//Also requires setting FirstVisibleRow if you want a specific row to be active in the top pane.
sheet.FirstVisibleRow = 5; // Example

# spire.xls c# excel Worksheets
## Set Active Pane in Split Worksheet
sheet.VerticalSplit = 2000;
sheet.HorizontalSplit = 3000;
//Set the bottom-right pane (index 3) as active. Panes are 0:TopLeft, 1:TopRight, 2:BottomLeft, 3:BottomRight
sheet.ActivePane = 3;

# spire.xls c# excel Worksheets
## Remove Panes from Worksheet
sheet.RemovePanes();
Console.WriteLine("All frozen or split panes removed from sheet '" + sheet.Name + "'.");

# spire.xls c# excel Worksheets
## Check if Worksheet is Password Protected
bool isProtected = sheet.IsPasswordProtected;
Console.WriteLine("Worksheet '" + sheet.Name + "' is password protected: " + isProtected);

# spire.xls c# excel Worksheets
## Set Worksheet Zoom Factor
//Set zoom factor to 75%
sheet.Zoom = 75;
Console.WriteLine("Zoom factor for sheet '" + sheet.Name + "' set to 75%.");

# spire.xls csharp data export
## export data while preserving data types

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Export DataTable with data type preservation
ExportTableOptions options = new ExportTableOptions();
options.ExportColumnNames = true;
options.KeepDataFormat = false;
options.KeepDataType = true;
options.RenameStrategy = RenameStrategy.Digit;

// Export data to data table
DataTable table = sheet.ExportDataTable(1, 1, sheet.LastDataRow, sheet.LastDataColumn, options);


# Remove Duplicated Rows in Excel
## Remove duplicate rows from an Excel worksheet using Spire.XLS
// Remove duplicated rows in the worksheet
sheet.RemoveDuplicates();

// Remove the duplicate rows within the specified range
// sheet.RemoveDuplicates(int startRow, int startColumn, int endRow, int endColumn);
// Remove the duplicated rows based on specific columns and headers
// sheet.RemoveDuplicates(int startRow, int startColumn, int endRow, int endColumn, boolean hasHeaders, int[] columnOffsets)


# Markdown to XLSX Conversion
## Convert Markdown files to Excel XLSX format using Spire.Xls library
// Create a new Workbook instance
Workbook workbook = new Workbook();

// Load content from a Markdown file into the workbook
workbook.LoadFromMarkdown(markdownFilePath);

// Save the workbook to an Excel file
workbook.SaveToFile(outputFileName, ExcelVersion.Version2016);

// Release the resources used by the workbook object
workbook.Dispose();


# Excel Shape Hyperlink
## Add hyperlink to shapes in Excel worksheet
// Get the reference to the first sheet in the workbook
Worksheet sheet = workbook.Worksheets[0];

// Get all the shapes in the sheet
PrstGeomShapeCollection prstGeomShapeType = sheet.PrstGeomShapes;

// Set the hyperlink for each shape
for (int i = 0; i < prstGeomShapeType.Count; i++)
{
    // Get the shape
    XlsPrstGeomShape shape = (XlsPrstGeomShape)prstGeomShapeType[i];

    // Set the hyperlink address
    shape.HyLink.Address = "https://www.e-iceblue.com/Download/download-excel-for-net-now.html";
}


# spire.xls csharp pivot table
## create pivot table group by value
// Get the reference to the first sheet in the workbook
Worksheet pivotSheet = workbook.Worksheets[0];

// Cast the first PivotTable in the PivotTables collection to an XlsPivotTable object.
XlsPivotTable pivot = (XlsPivotTable)pivotSheet.PivotTables[0];

// Retrieve the PivotField named "number" from the PivotTable and cast it to a PivotField object.
PivotField dateBaseField = pivot.PivotFields["number"] as PivotField;

// Create a group for the PivotField, starting at 3000, ending at 3800, with an interval of 1.
dateBaseField.CreateGroup(3000, 3800, 1);

// Recalculate the data in the PivotTable to reflect the changes made.
pivot.CalculateData();

# spire.xls csharp pivot table slicer
## create and configure slicers from pivot table

// Get pivot table collection
Spire.Xls.Collections.PivotTablesCollection pivotTables = worksheet.PivotTables;

//Add a PivotTable to the worksheet
CellRange dataRange = worksheet.Range["A1:C9"];
PivotCache cache = wb.PivotCaches.Add(dataRange);

//Cell to put the pivot table
Spire.Xls.PivotTable pt = worksheet.PivotTables.Add("TestPivotTable", worksheet.Range["A12"], cache);

//Drag the fields to the row area.
PivotField pf = pt.PivotFields["fruit"] as PivotField;
pf.Axis = AxisTypes.Row;
PivotField pf2 = pt.PivotFields["year"] as PivotField;
pf2.Axis = AxisTypes.Column;

//Drag the field to the data area.
pt.DataFields.Add(pt.PivotFields["amount"], "SUM of Count", SubtotalTypes.Sum);

//Set PivotTable style
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium10;

pt.CalculateData();

//Get slicer collection
XlsSlicerCollection slicers = worksheet.Slicers;

int index = slicers.Add(pt, "E12", 0);

XlsSlicer xlsSlicer = slicers[index];
xlsSlicer.Name = "xlsSlicer";
xlsSlicer.Width = 100;
xlsSlicer.Height = 120;
xlsSlicer.StyleType = SlicerStyleType.SlicerStyleLight2;
xlsSlicer.PositionLocked = true;

//Get SlicerCache object of current slicer
XlsSlicerCache slicerCache = xlsSlicer.SlicerCache;
slicerCache.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithNoData;

//Style setting
XlsSlicerCacheItemCollection slicerCacheItems = xlsSlicer.SlicerCache.SlicerCacheItems;
XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems[0];
xlsSlicerCacheItem.Selected = false;

XlsSlicerCollection slicers_2 = worksheet.Slicers;

IPivotField r1 = pt.PivotFields["year"];
int index_2 = slicers_2.Add(pt, "I12", r1);

XlsSlicer xlsSlicer_2 = slicers[index_2];
xlsSlicer_2.RowHeight = 40;
xlsSlicer_2.StyleType = SlicerStyleType.SlicerStyleLight3;
xlsSlicer_2.PositionLocked = false;

//Get SlicerCache object of current slicer
XlsSlicerCache slicerCache_2 = xlsSlicer_2.SlicerCache;
slicerCache_2.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithDataAtTop;

//Style setting
XlsSlicerCacheItemCollection slicerCacheItems_2 = xlsSlicer_2.SlicerCache.SlicerCacheItems;
XlsSlicerCacheItem xlsSlicerCacheItem_2 = slicerCacheItems_2[1];
xlsSlicerCacheItem_2.Selected = false;
pt.CalculateData();

# spire.xls csharp slicer
## create slicers from table in excel
// Get slicer collection
XlsSlicerCollection slicers = worksheet.Slicers;

//Create a table with the data from the specific cell range.
IListObject table = worksheet.ListObjects.Create("Super Table", worksheet.Range["A1:C9"]);

int count = 3;
int index = 0;
foreach (SlicerStyleType type in Enum.GetValues(typeof(SlicerStyleType)))
{
    count += 5;
    String range = "E" + count;
    index = slicers.Add(table, range.ToString(), 0);

    //Style setting
    XlsSlicer xlsSlicer = slicers[index];
    xlsSlicer.Name = "slicers_" + count;
    xlsSlicer.StyleType = type;
}

# spire.xls csharp slicer modification
## modify Excel slicer properties including style, caption, and filtering behavior
// Get the first worksheet in the workbook
Worksheet worksheet = wb.Worksheets[0];

// Get the slicer collection from the worksheet
XlsSlicerCollection slicers = worksheet.Slicers;

// Get the first slicer from the slicer collection
XlsSlicer xlsSlicer = slicers[0];

// Set the style of the slicer to a dark theme (style type 4)
xlsSlicer.StyleType = SlicerStyleType.SlicerStyleDark4;

// Change the caption (title) of the slicer
xlsSlicer.Caption = "Modified Slicer";

// Lock the position of the slicer to prevent it from being moved in the worksheet
xlsSlicer.PositionLocked = true;

// Get the collection of cache items associated with the slicer
XlsSlicerCacheItemCollection slicerCacheItems = xlsSlicer.SlicerCache.SlicerCacheItems;

// Get the first cache item in the collection
XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems[0];

// Deselect the cache item
xlsSlicerCacheItem.Selected = false;

// Get the display value of the cache item
string displayValue = xlsSlicerCacheItem.DisplayValue;

// Get the slicer cache associated with the slicer
XlsSlicerCache slicerCache = xlsSlicer.SlicerCache;

// Set the cross-filter type to show items even if they have no associated data
slicerCache.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithNoData;

# Spire.XLS C# Slicer Information Reader
## Read and extract information about slicers in an Excel file
// Load Excel file and get worksheet
Workbook wb = new Workbook();
wb.LoadFromFile("SlicerTemplate.xlsx");
Worksheet worksheet = wb.Worksheets[0];

// Get slicer collection
XlsSlicerCollection slicers = worksheet.Slicers;

// Iterate through each slicer
for (int i = 0; i < slicers.Count; i++)
{
    XlsSlicer xlsSlicer = slicers[i];
    
    // Get slicer properties
    string slicerName = xlsSlicer.Name;
    string slicerCaption = xlsSlicer.Caption;
    int numberOfColumns = xlsSlicer.NumberOfColumns;
    double columnWidth = xlsSlicer.ColumnWidth;
    double rowHeight = xlsSlicer.RowHeight;
    bool showCaption = xlsSlicer.ShowCaption;
    bool positionLocked = xlsSlicer.PositionLocked;
    double width = xlsSlicer.Width;
    double height = xlsSlicer.Height;

    // Get slicer cache
    XlsSlicerCache slicerCache = xlsSlicer.SlicerCache;
    
    // Get slicer cache properties
    string sourceName = slicerCache.SourceName;
    bool isTabular = slicerCache.IsTabular;
    string cacheName = slicerCache.Name;

    // Get slicer cache items
    XlsSlicerCacheItemCollection slicerCacheItems = slicerCache.SlicerCacheItems;
    XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems[1];
    
    // Get slicer cache item properties
    bool isSelected = xlsSlicerCacheItem.Selected;
}

// Clean up
wb.Dispose();

# spire.xls remove slicer
## remove slicers from excel worksheet
// Get the slicer collection from the worksheet
XlsSlicerCollection slicers = worksheet.Slicers;

// Example: Remove the first slicer in the collection 
// slicers.RemoveAt(0);

// Clear all slicers from the collection
slicers.Clear();



