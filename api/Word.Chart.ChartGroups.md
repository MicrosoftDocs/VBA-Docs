---
title: Chart.ChartGroups property (Word)
ms.prod: word
api_name:
- Word.Chart.ChartGroups
ms.assetid: ae4da68e-1e80-f683-b1ef-eb26aa753420
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartGroups property (Word)

Returns an object that represents either a single chart group or a collection of all the chart groups in the chart.


## Syntax

_expression_. `ChartGroups` (_Index_)

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The chart group number. If specified, a single  **[ChartGroup](Word.ChartGroup.md)** object is returned. If omitted, a **[ChartGroups](Word.ChartGroups.md)** object, which contains a collection of every **ChartGroup** object for that chart, is returned.|

## Example

The following example enables up and down bars for the first chart group of the first chart, and then sets their colors. You should run this example on a 2D line chart that contains two series that intersect at one or more data points.






```vb
With ActiveDocument.InlineShapes(1).Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]