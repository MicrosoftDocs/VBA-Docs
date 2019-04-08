---
title: Frameset.HeightType property (Word)
keywords: vbawd10.chm165806082
f1_keywords:
- vbawd10.chm165806082
ms.prod: word
api_name:
- Word.Frameset.HeightType
ms.assetid: 4d83e41c-d33c-a5b8-853c-e7581170ba4b
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.HeightType property (Word)

Returns or sets the width type for the specified frame on a frames page. Read/write  **WdFramesetSizeType**.


## Syntax

_expression_. `HeightType`

_expression_ Required. A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Example

This example sets the height of the first Frameset object in the specified frames page to 25 percent of the window height.


```vb
With ActiveDocument.ActiveWindow.Panes(1).Frameset 
 .HeightType = wdFramesetSizeTypePercent 
 .Height = 25 
End With
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]