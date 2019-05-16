---
title: TextFrame.WarpFormat property (Word)
keywords: vbawd10.chm162665366
f1_keywords:
- vbawd10.chm162665366
ms.prod: word
api_name:
- Word.TextFrame.WarpFormat
ms.assetid: 2ea707b9-0ed1-1196-2bf9-a32ae87d456a
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.WarpFormat property (Word)

Returns or sets the warp format (how the text is warped) for the specified text frame. Read/write [MsoWarpFormat](Office.MsoWarpFormat.md).


## Syntax

_expression_. `WarpFormat`

_expression_ A variable that represents a **[TextFrame](Word.TextFrame.md)** object.


## Example

The following code example shows how to set the warp format for the first shape on the active document.


```vb
ActiveDocument.Shapes(1).TextFrame.WarpFormat = msoWarpFormat15
```


## See also


[TextFrame Object](Word.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]