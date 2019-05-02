---
title: Shape.TopRelative property (Word)
keywords: vbawd10.chm161480905
f1_keywords:
- vbawd10.chm161480905
ms.prod: word
api_name:
- Word.Shape.TopRelative
ms.assetid: 5ae905f1-2e86-2aab-fe43-3be81f61847c
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.TopRelative property (Word)

Returns or sets a  **Single** that represents the relative top position of a shape. Read/write.


## Syntax

_expression_. `TopRelative`

 _expression_ An expression that returns a **[Shape](Word.Shape.md)** object.


## Remarks

Use this property with the  **[RelativeHorizontalPosition](Word.Shape.RelativeHorizontalPosition.md)** property. When set to **wdShapePositionRelativeNone** (-999999) (see the **[WdShapePositionRelative](Word.WdShapePositionRelative.md)** enumeration), this property should be ignored because the shape does not use percent positioning. The vertical position is solely determined by the **[Top](Word.Shape.Top.md)** property.


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]