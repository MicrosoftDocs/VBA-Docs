---
title: ShapeRange.PickUp method (Publisher)
keywords: vbapb10.chm2293795
f1_keywords:
- vbapb10.chm2293795
ms.prod: publisher
api_name:
- Publisher.ShapeRange.PickUp
ms.assetid: ebd62b6e-807a-821c-d8ea-ed9be289c433
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.PickUp method (Publisher)

Copies formatting from a shape or shape range so that it can be copied to another shape or shape range by using the **[Apply](Publisher.ShapeRange.Apply.md)** method.


## Syntax

_expression_.**PickUp**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Remarks

You must use the **PickUp** method to copy the formatting from a shape or shape range before using the **Apply** method; otherwise, an error occurs.


## Example

The following example copies the formatting from the first shape of the active publication to the second shape of the active publication.

```vb
With ActiveDocument.Pages(1) 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]