---
title: ShapeRange.TopRelative property (Word)
keywords: vbawd10.chm162857161
f1_keywords:
- vbawd10.chm162857161
ms.prod: word
api_name:
- Word.ShapeRange.TopRelative
ms.assetid: 6162d05b-0610-7a6b-0224-7bd6f658276b
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.TopRelative property (Word)

Returns or sets a  **Single** that represents the relative top position of a range of shapes. Read/write.


## Syntax

_expression_. `TopRelative`

 _expression_ An expression that returns a **[ShapeRange](Word.shaperange.md)** object.


## Remarks

Use this property with the  **[RelativeVerticalPosition](Word.ShapeRange.RelativeVerticalPosition.md)** property. When set to **wdShapePositionRelativeNone** (-999999) (see the **[WdShapePositionRelative](Word.WdShapePositionRelative.md)** enumeration), this property should be ignored because the shape does not use percent positioning. The vertical position is solely determined by the **[Top](Word.ShapeRange.Top.md)** property.


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]