---
title: Frameset.FrameDefaultURL property (Word)
keywords: vbawd10.chm165806116
f1_keywords:
- vbawd10.chm165806116
ms.prod: word
api_name:
- Word.Frameset.FrameDefaultURL
ms.assetid: 596f57d4-2514-8cd0-2d97-20618051fd6c
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.FrameDefaultURL property (Word)

Returns or sets the webpage or other document to be displayed in the specified frame when the frames page is opened. Read/write  **String**.


## Syntax

_expression_. `FrameDefaultURL`

_expression_ A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example sets the specified frame to display a local file named "Order.htm".


```vb
With ActiveDocument.ActiveWindow.ActivePane.Frameset 
 .FrameDefaultURL = "C:\Documents\Order.htm" 
 .FrameLinkToFile = True 
End With
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]