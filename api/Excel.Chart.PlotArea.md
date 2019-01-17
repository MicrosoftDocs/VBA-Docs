---
title: Chart.PlotArea property (Excel)
keywords: vbaxl10.chm149134
f1_keywords:
- vbaxl10.chm149134
ms.prod: excel
api_name:
- Excel.Chart.PlotArea
ms.assetid: f3c93a06-b398-a60a-d69d-8249652501eb
ms.date: 06/08/2017
localization_priority: Priority
---


# Chart.PlotArea property (Excel)

Returns a  **[PlotArea](Excel.PlotArea(object).md)** object that represents the plot area of a chart. Read-only.


## Syntax

_expression_. `PlotArea`

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Example

This example sets the color of the plot area interior of Chart1 to cyan.


```vb
Charts("Chart1").PlotArea.Interior.ColorIndex = 8
```


## See also


[Chart Object](Excel.Chart(object).md)

