---
title: Shape.LeftRelative property (Word)
keywords: vbawd10.chm161480904
f1_keywords:
- vbawd10.chm161480904
ms.prod: word
api_name:
- Word.Shape.LeftRelative
ms.assetid: a4fd7e18-9e04-8ea9-58d1-e2e816079ac3
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.LeftRelative property (Word)

Returns or sets a  **Single** that represents the relative left position of a shape. Read/write.


## Syntax

_expression_. `LeftRelative`

 _expression_ An expression that returns a **[Shape](Word.Shape.md)** object.


## Remarks

Use this property with the  **[RelativeHorizontalPosition](Word.Shape.RelativeHorizontalPosition.md)** property. When set to **wdShapePositionRelativeNone** (-999999) (see the **[WdShapePositionRelative](Word.WdShapePositionRelative.md)** enumeration), this property should be ignored because the shape does not use percent positioning. The horizontal position is solely determined by the **[Left](Word.Shape.Left.md)** property.


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]