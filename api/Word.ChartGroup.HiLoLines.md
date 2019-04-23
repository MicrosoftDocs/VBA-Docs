---
title: ChartGroup.HiLoLines property (Word)
keywords: vbawd10.chm263454740
f1_keywords:
- vbawd10.chm263454740
ms.prod: word
api_name:
- Word.ChartGroup.HiLoLines
ms.assetid: 452f4e5d-7ad8-76ad-5067-2df8a074d6d1
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.HiLoLines property (Word)

Returns the high-low lines for a series on a line chart. Read-only  **[HiLoLines](Word.HiLoLines.md)**.


## Syntax

_expression_.**HiLoLines**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Remarks

This property applies only to line charts. 


## Example

The following example enables high-low lines for chart group one of the first chart in the active document and then sets their line style, weight, and color. You should run the example on a 2D line chart that has three series of stock-quote-like data (high-low-close).


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasHiLoLines = True 
 With .HiLoLines.Border 
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