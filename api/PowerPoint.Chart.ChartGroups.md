---
title: Chart.ChartGroups method (PowerPoint)
keywords: vbapp10.chm684018
f1_keywords:
- vbapp10.chm684018
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.ChartGroups
ms.assetid: 23339025-6d5f-b51a-e2b4-6fc15a903cea
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartGroups method (PowerPoint)

Returns an object that represents either a single chart group or a collection of all the chart groups in the chart.


## Syntax

_expression_. `ChartGroups` (_Index_)

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Variant**|The chart group number. If specified, a single  **[ChartGroup](PowerPoint.ChartGroup.md)** object is returned. If omitted, a **[ChartGroups](PowerPoint.ChartGroups.md)** object is returned which contains a collection of every **ChartGroup** object for that chart.|

## Example

This example turns on up and down bars for the first chart group of the first chart, and then sets their colors. The example should be run on a 2D line chart containing two series that intersect at one or more data points.


```vb
With ActivePresentation.Slides(1).Shapes(1).Chart.ChartGroups(1)

    .HasUpDownBars = True

    .DownBars.Interior.ColorIndex = 3

    .UpBars.Interior.ColorIndex = 5

End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]