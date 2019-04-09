---
title: Paragraphs.Shading property (Word)
keywords: vbawd10.chm156762228
f1_keywords:
- vbawd10.chm156762228
ms.prod: word
api_name:
- Word.Paragraphs.Shading
ms.assetid: b732c59d-d861-00d8-fd00-6940449480a1
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.Shading property (Word)

Returns a  **Shading** object that refers to the shading formatting for the specified paragraphs.


## Syntax

_expression_. `Shading`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example applies yellow shading to the all paragraphs in the selection.


```vb
With Selection.Paragraphs.Shading 
 .Texture = wdTexture12Pt5Percent 
 .BackgroundPatternColorIndex = wdYellow 
 .ForegroundPatternColorIndex = wdBlack 
End With
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]