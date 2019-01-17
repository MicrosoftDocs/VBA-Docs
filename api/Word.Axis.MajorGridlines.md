---
title: Axis.MajorGridlines property (Word)
keywords: vbawd10.chm113049617
f1_keywords:
- vbawd10.chm113049617
ms.prod: word
api_name:
- Word.Axis.MajorGridlines
ms.assetid: 90e0d7c0-add7-9a34-8706-aaf33f799441
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MajorGridlines property (Word)

Returns the major gridlines for the specified axis. Read-only  **[Gridlines](Word.GridLines.md)**.


## Syntax

 _expression_. `MajorGridlines`

 _expression_ A variable that represents an '[Axis](Word.Axis.md)' object.


## Remarks

Only axes in the primary axis group can have gridlines.


## Example

The following example sets the color of the major gridlines for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 If .HasMajorGridlines Then 
 ' Set the color to blue. 
 .MajorGridlines.Border.ColorIndex = 5 
 End If 
 End With 
 End If 
End With 

```


## See also


[Axis Object](Word.Axis.md)

