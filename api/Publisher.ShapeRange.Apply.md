---
title: ShapeRange.Apply method (Publisher)
keywords: vbapb10.chm2293776
f1_keywords:
- vbapb10.chm2293776
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Apply
ms.assetid: 3531d0aa-479e-2d50-5e1e-a35f7c1e7ba6
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Apply method (Publisher)

Applies formatting copied from another shape or shape range by using the **[PickUp](Publisher.ShapeRange.PickUp.md)** method.


## Syntax

_expression_.**Apply**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Return value

Nothing


## Remarks

If you do not first use the **PickUp** method to copy the formatting from another shape, an error occurs.


## Example

The following example copies the formatting from the first shape of the active publication to the second shape of the active publication.

```vb
With ActiveDocument.Pages(1) 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]