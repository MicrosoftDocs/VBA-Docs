---
title: ShapeRange.HeightRelative property (Word)
keywords: vbawd10.chm162857163
f1_keywords:
- vbawd10.chm162857163
ms.prod: word
api_name:
- Word.ShapeRange.HeightRelative
ms.assetid: f0414af1-f09a-475d-5e96-bfe2ceab8901
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.HeightRelative property (Word)

Returns or sets a  **Single** that represents the percentage of the target shape to which the range of shapes is sized. Read/write.


## Syntax

_expression_. `HeightRelative`

 _expression_ An expression that returns a **[ShapeRange](Word.shaperange.md)** object.


## Remarks

Use this property with the  **[RelativeVerticalSize](Word.ShapeRange.RelativeVerticalSize.md)** property. When set to **wdShapeSizeRelativeNone** (-999999) (see the **[WdShapeSizeRelative](Word.WdShapeSizeRelative.md)** enumeration), this property should be ignored because the shape does not use percent sizing. The height is solely determined by the **[Height](Word.ShapeRange.Height.md)** property.


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]