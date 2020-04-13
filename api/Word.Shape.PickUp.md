---
title: Shape.PickUp method (Word)
keywords: vbawd10.chm161480721
f1_keywords:
- vbawd10.chm161480721
ms.prod: word
api_name:
- Word.Shape.PickUp
ms.assetid: 9ccc7644-6186-d827-3dbe-db7dd3ccb4b6
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.PickUp method (Word)

Copies the formatting of the specified shape.


## Syntax

_expression_.**PickUp**

_expression_ Required. A variable that represents a **[Shape](Word.Shape.md)** object.


## Remarks

Use the **[Apply](Word.Shape.Apply.md)** method to apply the copied formatting to another shape.


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


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]