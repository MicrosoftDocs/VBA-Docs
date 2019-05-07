---
title: Frameset.FrameName property (Word)
keywords: vbawd10.chm165806114
f1_keywords:
- vbawd10.chm165806114
ms.prod: word
api_name:
- Word.Frameset.FrameName
ms.assetid: f0b22dfe-3d12-0f75-1af2-23467b83a4ad
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.FrameName property (Word)

Returns or sets the name of the specified frame on a frames page. Read/write  **String**.


## Syntax

_expression_. `FrameName`

_expression_ A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example sets the name of the specified frame to "BottomFrame".


```vb
ActiveWindow.Document.Frameset _ 
 .ChildFramesetItem(3).FrameName = "BottomFrame"
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]