---
title: Chart.GetChartElement method (Word)
keywords: vbawd10.chm79365505
f1_keywords:
- vbawd10.chm79365505
ms.prod: word
api_name:
- Word.Chart.GetChartElement
ms.assetid: e9ebb101-1625-9a6a-1da9-dbb02c49f01c
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.GetChartElement method (Word)

Returns information about the chart element at the specified x-coordinate and y-coordinate. 


## Syntax

_expression_.**GetChartElement** (_x_, _y_, _ElementID_, _Arg1_, _Arg2_)

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _x_|Required| **Long**|The x-coordinate of the chart element.|
| _y_|Required| **Long**|The y-coordinate of the chart element.|
| _ElementID_|Required| **Long**|When the method returns, this argument contains the **[XlChartItem](Word.xlchartitem.md)** value of the chart element at the specified coordinates. For more information, see the Remarks section.|
| _Arg1_|Required| **Long**|When the method returns, this argument contains information related to the chart element. For more information, see the Remarks section.|
| _Arg2_|Required| **Long**|When the method returns, this argument contains information related to the chart element. For more information, see the Remarks section.|

## Remarks

This method is unusual in that you specify values for only the first two arguments. Microsoft Word fills in the other arguments, and your code should examine those values when the method returns.

The value of ElementID after the method returns determines whether Arg1 and Arg2 contain any information, as shown in the following table.



|**ElementID Constant**|**Constant Value **|**Arg1**|**Arg2**|
|:-----|:-----|:-----|:-----|
| **xlAxis**|21|AxisIndex|AxisType|
| **xlAxisTitle**|17|AxisIndex|AxisType|
| **xlDisplayUnitLabel**|30|AxisIndex|AxisType|
| **xlMajorGridlines**|15|AxisIndex|AxisType|
| **xlMinorGridlines**|16|AxisIndex|AxisType|
| **xlPivotChartDropZone**|32|DropZoneType|None|
| **xlPivotChartFieldButton**|31|DropZoneType|None|
| **xlDownBars**|20|GroupIndex|None|
| **xlDropLines**|26|GroupIndex|None|
| **xlHiLoLines**|25|GroupIndex|None|
| **xlRadarAxisLabels**|27|GroupIndex|None|
| **xlSeriesLines**|22|GroupIndex|None|
| **xlUpBars**|18|GroupIndex|None|
| **xlChartArea**|2|None|None|
| **xlChartTitle**|4|None|None|
| **xlCorners**|6|None|None|
| **xlDataTable**|7|None|None|
| **xlFloor**|23|None|None|
| **xlLeaderLines**|29|None|None|
| **xlLegend**|24|None|None|
| **xlNothing**|28|None|None|
| **xlPlotArea**|19|None|None|
| **xlWalls**|5|None|None|
| **xlDataLabel**|7|SeriesIndex|PointIndex|
| **xlErrorBars**|9|SeriesIndex|None|
| **xlLegendEntry**|12|SeriesIndex|None|
| **xlLegendKey**|13|SeriesIndex|None|
| **xlSeries**|3|SeriesIndex|PointIndex|
| **xlShape**|14|ShapeIndex|None|
| **xlTrendline**|8|SeriesIndex|TrendLineIndex|
| **xlXErrorBars**|10|SeriesIndex|None|
| **xlYErrorBars**|11|SeriesIndex|None|

The following table describes the meaning of Arg1 and Arg2 after the method returns.



|**Argument**|**Description**|
|:-----|:-----|
|AxisIndex|Specifies whether the axis is primary or secondary. Can be one of the following  **[XlAxisGroup](Word.xlaxisgroup.md)** constants: **xlPrimary** or **xlSecondary**.|
|AxisType|Specifies the axis type. Can be one of the following  **[XlAxisType](Word.xlaxistype.md)** constants: **xlCategory**, **xlSeriesAxis**, or **xlValue**.|
|DropZoneType|Specifies the drop zone type: column, data, page, or row field. Can be one of the following  **[XlPivotFieldOrientation](Word.xlpivotfieldorientation.md)** constants: **xlColumnField**, **xlDataField**, **xlPageField**, or **xlRowField**. The column and row field constants specify the series and category fields, respectively.|
|GroupIndex|Specifies the offset within the **[ChartGroups](Word.ChartGroups.md)** collection for a specific chart group.|
|PointIndex|Specifies the offset within the **[Points](Word.Points.md)** collection for a specific point within a series. A value of 1 indicates that all data points are selected.|
|SeriesIndex|Specifies the offset within the **[Series](Word.Series.md)** collection for a specific series.|
|ShapeIndex|Specifies the offset within the **[Shapes](Word.shapes.md)** collection for a specific shape.|
|TrendlineIndex|Specifies the offset within the **[Trendlines](Word.Trendlines.md)** collection for a specific trendline within a series.|

## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]