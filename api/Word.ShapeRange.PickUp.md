---
title: ShapeRange.PickUp method (Word)
keywords: vbawd10.chm162856980
f1_keywords:
- vbawd10.chm162856980
ms.prod: word
api_name:
- Word.ShapeRange.PickUp
ms.assetid: 6074168d-5cb2-2f86-fca4-c609dd2333f8
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.PickUp method (Word)

Copies the formatting of the specified shape.


## Syntax

_expression_.**PickUp**

_expression_ Required. A variable that represents a **[ShapeRange](Word.shaperange.md)** object.


## Remarks

Use the **[Apply](Word.ShapeRange.Apply.md)** method to apply the copied formatting to another shape.


## Example

This example copies the formatting of shape one on _myDocument_ and then applies the copied formatting to shape two.


```vb
Set myDocument = ActiveDocument 
With myDocument 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With
```


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]