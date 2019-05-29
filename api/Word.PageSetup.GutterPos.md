---
title: PageSetup.GutterPos property (Word)
keywords: vbawd10.chm158401734
f1_keywords:
- vbawd10.chm158401734
ms.prod: word
api_name:
- Word.PageSetup.GutterPos
ms.assetid: 71027b04-e01b-e826-c0ae-39ca3c33182a
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.GutterPos property (Word)

Returns or sets on which side the gutter appears in a document. Read/write  **WdGutterStyle**.


## Syntax

_expression_. `GutterPos`

_expression_ Required. A variable that represents a **[PageSetup](Word.PageSetup.md)** object.


## Example

This example sets the gutter to appear on the right side of the document.


```vb
ActiveDocument.PageSetup.GutterPos = wdGutterPosRight
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]