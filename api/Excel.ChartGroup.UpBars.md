---
title: ChartGroup.UpBars property (Excel)
keywords: vbaxl10.chm568092
f1_keywords:
- vbaxl10.chm568092
ms.prod: excel
api_name:
- Excel.ChartGroup.UpBars
ms.assetid: d97b23bd-4c51-2384-a5f3-7cc067d3d6fa
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.UpBars property (Excel)

Returns an **[UpBars](Excel.UpBars(object).md)** object that represents the up bars on a line chart. Applies only to line charts. Read-only.


## Syntax

_expression_.**UpBars**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example turns on up and down bars for chart group one on Chart1 and then sets their colors. The example should be run on a 2D line chart containing two series that cross each other at one or more data points.

```vb
With Charts("Chart1").ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]