---
title: HiLoLines object (Word)
keywords: vbawd10.chm3601
f1_keywords:
- vbawd10.chm3601
ms.prod: word
api_name:
- Word.HiLoLines
ms.assetid: 9f1ed891-7e95-8dd0-745a-ce28555284a9
ms.date: 06/08/2017
localization_priority: Normal
---


# HiLoLines object (Word)

Represents the high-low lines in a chart group.


## Remarks

 High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2D line groups can have high-low lines. This object is not a collection. There is no object that represents a single high-low line; you either enable high-low lines for all points in a chart group or disable them.

If the **[HasHiLoLines](Word.ChartGroup.HasHiLoLines.md)** property is **False**, most properties of the **HiLoLines** object are disabled.


## Example

Use the **[HiLoLines](Word.ChartGroup.HiLoLines.md)** property to return the **HiLoLines** object. The following example uses the **HasHiLowLines** property to add high-low lines to the first chart (the chart must be a line chart) in the active document. The example then makes the high-low lines blue.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With Chart.ChartGroups(1) 
 .HasHighLowLines = True 
 .HiLoLines.Border.Color = RGB(0, 0, 255) 
 End With 
 End If 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]