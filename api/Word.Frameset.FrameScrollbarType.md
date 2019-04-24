---
title: Frameset.FrameScrollbarType property (Word)
keywords: vbawd10.chm165806110
f1_keywords:
- vbawd10.chm165806110
ms.prod: word
api_name:
- Word.Frameset.FrameScrollbarType
ms.assetid: dacd6394-872e-beac-85dc-575234f9ce29
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.FrameScrollbarType property (Word)

Returns or sets when scroll bars are available for the specified frame when viewing its frames page in a web browser. Read/write  **WdScrollbarType**.


## Syntax

_expression_. `FrameScrollbarType`

_expression_ Required. A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Remarks

For more information on creating frames pages, see [Creating frames pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example makes scroll bars always available for the specified frame, regardless of whether the contents of the frame require scrolling.


```vb
With ActiveDocument.ActiveWindow.ActivePane.Frameset 
 .FrameDefaultURL = "C:\Documents\Order.htm" 
 .FrameScrollBarType = wdScrollBarTypeYes 
End With
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]