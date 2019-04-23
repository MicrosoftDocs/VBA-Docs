---
title: Gridlines object (Word)
keywords: vbawd10.chm175
f1_keywords:
- vbawd10.chm175
ms.prod: word
api_name:
- Word.GridLines
ms.assetid: 9dc77c2a-854f-63c0-4648-b7802fb6d9a2
ms.date: 06/08/2017
localization_priority: Normal
---


# Gridlines object (Word)

Represents major or minor gridlines on a chart axis.


## Remarks

 Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers. This object is not a collection. There is no object that represents a single gridline; you either enable all gridlines for an axis or disable all of them.

Use the  **[MajorGridlines](Word.Axis.MajorGridlines.md)** property to return the **GridLines** object that represents the major gridlines for the axis. Use the **[MinorGridlines](Word.Axis.MinorGridlines.md)** property to return the **GridLines** object that represents the minor gridlines. It is possible to return both major and minor gridlines at the same time.


## Example

The following example enables major gridlines for the category axis of the first chart in the active document and then formats the gridlines to be blue dashed lines.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory) 
 .HasMajorGridlines = True 
 .MajorGridlines.Border.Color = RGB(0, 0, 255) 
 .MajorGridlines.Border.LineStyle = xlDash 
 End With 
 End If 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]