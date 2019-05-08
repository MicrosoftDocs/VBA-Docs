---
title: ChartGroup.DownBars property (Excel)
keywords: vbaxl10.chm568075
f1_keywords:
- vbaxl10.chm568075
ms.prod: excel
api_name:
- Excel.ChartGroup.DownBars
ms.assetid: dd8ae50c-0105-9645-467d-7eb07b56c95e
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.DownBars property (Excel)

Returns a **[DownBars](Excel.DownBars(object).md)** object that represents the down bars on a line chart. Applies only to line charts. Read-only.


## Syntax

_expression_.**DownBars**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example turns on up bars and down bars for chart group one on Chart1 and then sets their colors. The example should be run on a 2D line chart that has two series that cross each other at one or more data points.

```vb
With Charts("Chart1").ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]