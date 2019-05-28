---
title: Options.DefaultBorderColorIndex property (Word)
keywords: vbawd10.chm162988369
f1_keywords:
- vbawd10.chm162988369
ms.prod: word
api_name:
- Word.Options.DefaultBorderColorIndex
ms.assetid: 8d430be3-b27e-7650-0c23-87436f088a0b
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DefaultBorderColorIndex property (Word)

Returns or sets the default line color for borders. Read/write  **WdColorIndex**.


## Syntax

_expression_. `DefaultBorderColorIndex`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example changes the default line color and style for borders and then applies a border around the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Borders.Enable = True 
With Options 
 .DefaultBorderColorIndex = wdRed 
 .DefaultBorderLineStyle = wdLineStyleDouble 
End With
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]