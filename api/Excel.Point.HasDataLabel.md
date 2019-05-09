---
title: Point.HasDataLabel property (Excel)
keywords: vbaxl10.chm576081
f1_keywords:
- vbaxl10.chm576081
ms.prod: excel
api_name:
- Excel.Point.HasDataLabel
ms.assetid: 924f70a0-fdeb-e155-c857-55e0dfb7ca60
ms.date: 05/09/2019
localization_priority: Normal
---


# Point.HasDataLabel property (Excel)

**True** if the point has a data label. Read/write **Boolean**.


## Syntax

_expression_.**HasDataLabel**

_expression_ A variable that represents a **[Point](Excel.Point(object).md)** object.


## Example

This example turns on the data label for point seven in series three on Chart1, and then it sets the data label color to blue.

```vb
With Charts("Chart1").SeriesCollection(3).Points(7) 
 .HasDataLabel = True 
 .ApplyDataLabels Type:=xlValue 
 .DataLabel.Font.ColorIndex = 5 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]