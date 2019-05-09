---
title: ChartGroup.HasUpDownBars property (Excel)
keywords: vbaxl10.chm568083
f1_keywords:
- vbaxl10.chm568083
ms.prod: excel
api_name:
- Excel.ChartGroup.HasUpDownBars
ms.assetid: 891f305c-521c-3ec5-3e88-886e1dbdaea2
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.HasUpDownBars property (Excel)

**True** if a line chart has up and down bars. Applies only to line charts. Read/write **Boolean**.


## Syntax

_expression_.**HasUpDownBars**

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