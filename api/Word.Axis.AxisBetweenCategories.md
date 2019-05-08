---
title: Axis.AxisBetweenCategories property (Word)
keywords: vbawd10.chm113049600
f1_keywords:
- vbawd10.chm113049600
ms.prod: word
api_name:
- Word.Axis.AxisBetweenCategories
ms.assetid: b99e83a2-5540-e69d-402c-224612f8e568
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.AxisBetweenCategories property (Word)

 **True** if the value axis crosses the category axis between categories. Read/write **Boolean**.


## Syntax

_expression_.**AxisBetweenCategories**

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

This property applies only to category axes, and it does not apply to 3D charts.


## Example

The following example causes the value axis for the first chart in the active document to cross the category axis between categories.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlCategory). _ 
 AxisBetweenCategories = True 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]