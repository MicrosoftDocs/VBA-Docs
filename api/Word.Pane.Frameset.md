---
title: Pane.Frameset property (Word)
keywords: vbawd10.chm157286418
f1_keywords:
- vbawd10.chm157286418
ms.prod: word
api_name:
- Word.Pane.Frameset
ms.assetid: 6bab63ae-aa83-e2b8-9b92-e472c2433246
ms.date: 06/08/2017
localization_priority: Normal
---


# Pane.Frameset property (Word)

Returns a  **[Frameset](Word.Frameset.md)** object that represents an entire frames page or a single frame on a frames page. Read-only.


## Syntax

_expression_. `Frameset`

_expression_ A variable that represents a '[Pane](Word.Pane.md)' object.


## Remarks

For more information on creating frames pages, see [Creating frames pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).




## Example

This example adds a new frame to the immediate right of the specified frame.


```vb
ActiveDocument.ActiveWindow.ActivePane.Frameset _ 
 .AddNewFrame wdFramesetNewRight
```


## See also


[Pane Object](Word.Pane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]