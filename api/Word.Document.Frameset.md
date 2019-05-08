---
title: Document.Frameset property (Word)
keywords: vbawd10.chm158007623
f1_keywords:
- vbawd10.chm158007623
ms.prod: word
api_name:
- Word.Document.Frameset
ms.assetid: 40079f4f-be1d-c8dd-5536-ccb5f570bde9
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Frameset property (Word)

Returns a  **[Frameset](Word.Frameset.md)** object that represents an entire frames page or a single frame on a frames page. Read-only.


## Syntax

_expression_. `Frameset`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For more information on creating frames pages, see [Creating frames pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example sets the color of frame borders in the specified frames page to tan.


```vb
With ActiveWindow.Document.Frameset 
 .FramesetBorderColor = wdColorTan 
 .FramesetBorderWidth = 6 
End With
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]