---
title: Chart.ChartWizard method (Word)
keywords: vbawd10.chm79364162
f1_keywords:
- vbawd10.chm79364162
ms.prod: word
api_name:
- Word.Chart.ChartWizard
ms.assetid: 5c4c4cb1-3ef7-e3c3-d441-6f92cb8e7771
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartWizard method (Word)

Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.

## Syntax

_expression_.**ChartWizard** (**_Source_**, **_Gallery_**, **_Format_**, **_PlotBy_**, **_CategoryLabels_**, **_SeriesLabels_**, **_HasLegend_**, **_Title_**, **_CategoryTitle_**, **_ValueTitle_**, **_ExtraTitle_**)

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Optional|**Variant**|The range that contains the source data for the new chart. If this argument is omitted, Word edits the active chart sheet or the selected chart on the active worksheet.|
| _Gallery_|Optional|**Variant**|One of the **[XlChartType](Excel.XlChartType.md)** constants that specifies the chart type.|
| _Format_|Optional|**Variant**|The option number for the built-in autoformats. Can be a number from 1 through 10, depending on the gallery type. If this argument is omitted, Word chooses a default value based on the gallery type and data source.|
| _PlotBy_|Optional|**Variant**|Specifies whether the data for each series is in rows or columns. Can be one of the following **[XlRowCol](Word.xlrowcol.md)** constants: **xlRows** or **xlColumns**.|
| _CategoryLabels_|Optional|**Variant**|An integer that specifies the number of rows or columns within the source range that contain category labels. Allowed values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.|
| _SeriesLabels_|Optional|**Variant**|An integer that specifies the number of rows or columns within the source range that contain series labels. Allowed values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.|
| _HasLegend_|Optional|**Variant**|**True** to include a legend.|
| _Title_|Optional|**Variant**|The chart title text.|
| _CategoryTitle_|Optional|**Variant**|The category axis title text.|
| _ValueTitle_|Optional|**Variant**|The value axis title text.|
| _ExtraTitle_|Optional|**Variant**|The series axis title for 3D charts or the second value axis title for 2D charts.|

<br/>

## Remarks

If the Source parameter is omitted and the selection is not a chart on the active document, this method fails and an error occurs.

## Example

The following example reformats the first chart as a line chart, adds a legend, and adds category and value axis titles.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 .ChartWizard _ 
 Gallery:=xlLine, _ 
 HasLegend:=True, _ 
 CategoryTitle:="Year", _ 
 ValueTitle:="Sales" 
End With
```


## See also

- [Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]