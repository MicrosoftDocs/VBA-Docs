---
title: Frameset.WidthType property (Word)
keywords: vbawd10.chm165806081
f1_keywords:
- vbawd10.chm165806081
ms.prod: word
api_name:
- Word.Frameset.WidthType
ms.assetid: a5e998bc-317a-dc62-a139-4e5ada8a4866
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.WidthType property (Word)

Returns or sets the width type for the specified  **Frameset** object. Read/write **WdFramesetSizeType**.


## Syntax

 _expression_. `WidthType`

 _expression_ Required. A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Example

This example sets the width of the first  **Frameset** object in the active document to 25% of the window width.


```vb
With ActiveDocument.ActiveWindow.Panes(1).Frameset 
 .WidthType = wdFramesetSizeTypePercent 
 .Width = 25 
End With
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]