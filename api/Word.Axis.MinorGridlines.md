---
title: Axis.MinorGridlines property (Word)
keywords: vbawd10.chm113049636
f1_keywords:
- vbawd10.chm113049636
ms.prod: word
api_name:
- Word.Axis.MinorGridlines
ms.assetid: b234c5ca-0381-6834-b2f9-fae3048a2fbf
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MinorGridlines property (Word)

Returns the minor gridlines for the specified axis. Read-only  **[Gridlines](Word.GridLines.md)**.


## Syntax

_expression_.**MinorGridlines**

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

Only axes in the primary axis group can have gridlines.


## Example

The following example sets the color of the minor gridlines for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 If .HasMinorGridlines Then 
 ' Set the color to blue. 
 .MinorGridlines.Border.ColorIndex = 5 
 End If 
 End With 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]