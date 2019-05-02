---
title: Shape.HeightRelative property (Word)
keywords: vbawd10.chm161480907
f1_keywords:
- vbawd10.chm161480907
ms.prod: word
api_name:
- Word.Shape.HeightRelative
ms.assetid: 24a52ebf-1071-a2e4-8222-9b17d295e653
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.HeightRelative property (Word)

Returns or sets a  **Single** that represents the percentage of the relative height of a shape. Read/write.


## Syntax

_expression_. `HeightRelative`

 _expression_ An expression that returns a **[Shape](Word.Shape.md)** object.


## Remarks

Use this property with the  **[RelativeVerticalSize](Word.Shape.RelativeVerticalSize.md)** property. When set to **wdShapeSizeRelativeNone** (-999999) (see the **[WdShapeSizeRelative](Word.WdShapeSizeRelative.md)** enumeration), this property should be ignored because the shape does not use percent sizing. The height is solely determined by the **[Height](Word.Shape.Height.md)** property.


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]