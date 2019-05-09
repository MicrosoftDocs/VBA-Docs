---
title: ChartGroup.HasDropLines property (Excel)
keywords: vbaxl10.chm568079
f1_keywords:
- vbaxl10.chm568079
ms.prod: excel
api_name:
- Excel.ChartGroup.HasDropLines
ms.assetid: cc0d188d-51ba-951d-7063-10820e5e4a42
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.HasDropLines property (Excel)

**True** if the line chart or area chart has drop lines. Applies only to line and area charts. Read/write **Boolean**.


## Syntax

_expression_.**HasDropLines**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example turns on drop lines for chart group one on Chart1, and then sets their line style, weight, and color. The example should be run on a 2D line chart that has one series.

```vb
With Charts("Chart1").ChartGroups(1) 
 .HasDropLines = True 
 With .DropLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]