---
title: ChartGroup.DropLines property (Word)
keywords: vbawd10.chm263454725
f1_keywords:
- vbawd10.chm263454725
ms.prod: word
api_name:
- Word.ChartGroup.DropLines
ms.assetid: eebe1c74-5682-4680-56d2-f0190fec5950
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.DropLines property (Word)

Returns the drop lines for a series on a line chart or area chart. Read-only  **[DropLines](Word.DropLines.md)**.


## Syntax

_expression_.**DropLines**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Remarks

This property applies only to line charts or area charts. 


## Example

The following example enables drop lines for chart group one of the first chart in the active document and then sets their line style, weight, and color. You should run the example on a 2D line chart that has one series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With Chart.ChartGroups(1) 
 .HasDropLines = True 
 With .DropLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
 End With 
 End If 
End With
```


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]