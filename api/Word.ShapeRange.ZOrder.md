---
title: ShapeRange.ZOrder method (Word)
keywords: vbawd10.chm162856988
f1_keywords:
- vbawd10.chm162856988
ms.prod: word
api_name:
- Word.ShapeRange.ZOrder
ms.assetid: 7f9a1a08-ac21-8866-9bf7-6a850200e2fd
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.ZOrder method (Word)

Moves the specified shape range in front of or behind other shapes in the collection (that is, changes the shape range's position in the z-order).


## Syntax

_expression_.**ZOrder** (_ZOrderCmd_)

 _expression_ An expression that returns a **[ShapeRange](Word.shaperange.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required| **MsoZOrderCmd**|Specifies where to move the specified shape range relative to the other shapes.|

## Return value

Nothing


## Remarks

Use the **[ZOrderPosition](Word.ShapeRange.ZOrderPosition.md)** property to determine a shape range's current position in the z-order.


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]