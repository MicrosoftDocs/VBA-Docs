---
title: Frameset.FrameResizable property (Word)
keywords: vbawd10.chm165806111
f1_keywords:
- vbawd10.chm165806111
ms.prod: word
api_name:
- Word.Frameset.FrameResizable
ms.assetid: 5a373e57-3193-c2a3-52b6-42702237f6c3
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.FrameResizable property (Word)

 **True** if the user can resize the specified frame when the frames page is viewed in a web browser. Read/write **Boolean**.


## Syntax

_expression_. `FrameResizable`

_expression_ A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example sets the specified frame to be resizable when viewed in a web browser.


```vb
With ActiveDocument.ActiveWindow.ActivePane.Frameset 
 .FrameDefaultURL = "C:\Documents\Order.htm" 
 .FrameResizable = True 
End With
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]