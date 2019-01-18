---
title: Chart.ChartWizard method (Excel)
keywords: vbaxl10.chm149090
f1_keywords:
- vbaxl10.chm149090
ms.prod: excel
api_name:
- Excel.Chart.ChartWizard
ms.assetid: c47588d9-6969-d6bb-cbbc-4941198d78b4
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartWizard method (Excel)

Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.


## Syntax

_expression_. `ChartWizard`( `_Source_` , `_Gallery_` , `_Format_` , `_PlotBy_` , `_CategoryLabels_` , `_SeriesLabels_` , `_HasLegend_` , `_Title_` , `_CategoryTitle_` , `_ValueTitle_` , `_ExtraTitle_` )

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Optional| **Variant**|The range that contains the source data for the new chart. If this argument is omitted, Microsoft Excel edits the active chart sheet or the selected chart on the active worksheet.|
| _Gallery_|Optional| **Variant**|One of the constants of  **[xlChartType](Excel.XlChartType.md)** specifying the chart type.|
| _Format_|Optional| **Variant**|The option number for the built-in autoformats. Can be a number from 1 through 10, depending on the gallery type. If this argument is omitted, Microsoft Excel chooses a default value based on the gallery type and data source.|
| _PlotBy_|Optional| **Variant**|Specifies whether the data for each series is in rows or columns. Can be one of the following  **[xlRowCol](Excel.XlRowCol.md)** constants: **xlRows** or **xlColumns**.|
| _CategoryLabels_|Optional| **Variant**|An integer specifying the number of rows or columns within the source range that contain category labels. Legal values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.|
| _SeriesLabels_|Optional| **Variant**|An integer specifying the number of rows or columns within the source range that contain series labels. Legal values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.|
| _HasLegend_|Optional| **Variant**| **True** to include a legend.|
| _Title_|Optional| **Variant**|The chart title text.|
| _CategoryTitle_|Optional| **Variant**|The category axis title text.|
| _ValueTitle_|Optional| **Variant**|The value axis title text.|
| _ExtraTitle_|Optional| **Variant**| The series axis title for 3-D charts or the second value axis title for 2-D charts.|

## Remarks

If  _Source_ is omitted and either the selection isn't an embedded chart on the active worksheet or the active sheet isn't an existing chart, this method fails and an error occurs.


## Example

This example reformats Chart1 as a line chart, adds a legend, and adds category and value axis titles.


```vb
Charts("Chart1").ChartWizard _ 
 Gallery:=xlLine, _ 
 HasLegend:=True, CategoryTitle:="Year", ValueTitle:="Sales"
```


## See also


[Chart Object](Excel.Chart(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]