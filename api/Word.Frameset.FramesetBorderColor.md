---
title: Frameset.FramesetBorderColor property (Word)
keywords: vbawd10.chm165806101
f1_keywords:
- vbawd10.chm165806101
ms.prod: word
api_name:
- Word.Frameset.FramesetBorderColor
ms.assetid: c47a7b7e-17e0-1741-fd1c-22cde123b42f
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.FramesetBorderColor property (Word)

Returns or sets the color of the frame borders on the specified frames page. Read/write.


## Syntax

_expression_. `FramesetBorderColor`

_expression_ Required. A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Remarks

This property can be any of the **WdColor** constants or a value returned by Visual Basic's **RGB** function. For more information on creating frames pages, see [Creating frames pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example sets the color of frame borders in the specified frames page to tan.


```vb
With ActiveWindow.Document.Frameset 
 .FramesetBorderColor = wdColorTan 
 .FramesetBorderWidth = 6 
End With
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]