---
title: Chart.ChartType property (Excel)
keywords: vbaxl10.chm149149
f1_keywords:
- vbaxl10.chm149149
ms.prod: excel
api_name:
- Excel.Chart.ChartType
ms.assetid: 532a2988-babf-b51a-7548-2f11f94c82a6
ms.date: 06/08/2017
localization_priority: Priority
---


# Chart.ChartType property (Excel)

Returns or sets the chart type. Read/write  **[xlChartType](Excel.XlChartType.md)**.


## Syntax

_expression_. `ChartType`

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Remarks

Some chart types aren't available for PivotChart reports.


## Example

This example sets the bubble size in chart group one to 200% of the default size if the chart is a 2D bubble chart.


```vb
With Worksheets(1).ChartObjects(1).Chart 
 If .ChartType = xlBubble Then 
 .ChartGroups(1).BubbleScale = 200 
 End If 
End With
```


## See also


[Chart Object](Excel.Chart(object).md)

