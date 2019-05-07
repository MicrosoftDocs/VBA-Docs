---
title: Axis.CategoryNames property (Word)
keywords: vbawd10.chm113049604
f1_keywords:
- vbawd10.chm113049604
ms.prod: word
api_name:
- Word.Axis.CategoryNames
ms.assetid: 12cb3d4e-1460-3849-5ce0-df9f0648d418
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.CategoryNames property (Word)

Returns or sets all the category names as a text array for the specified axis. Read/write  **Variant**.


## Syntax

_expression_.**CategoryNames**

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Example

The following example uses an array to set individual category names for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlCategory).CategoryNames = _ 
 Array ("1985", "1986", "1987", "1988", "1989") 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]