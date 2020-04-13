---
title: Shading.ForegroundPatternColorIndex property (Word)
keywords: vbawd10.chm154796033
f1_keywords:
- vbawd10.chm154796033
ms.prod: word
api_name:
- Word.Shading.ForegroundPatternColorIndex
ms.assetid: 9a6e7647-b034-7ae3-55ca-9d0e1956b76f
ms.date: 06/08/2017
localization_priority: Normal
---


# Shading.ForegroundPatternColorIndex property (Word)

Returns or sets the color that's applied to the foreground of the **Shading** object. This color is applied to the dots and lines in the shading pattern. Read/write **WdColorIndex**.


## Syntax

_expression_. `ForegroundPatternColorIndex`

_expression_ Required. A variable that represents a '[Shading](Word.Shading.md)' object.


## Example

This example applies shading with different foreground and background colors to the selection.


```vb
With Selection.Shading 
 .Texture = wdTexture30Percent 
 .ForegroundPatternColorIndex = wdBlue 
 .BackgroundPatternColorIndex = wdYellow 
End With
```


## See also


[Shading Object](Word.Shading.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]