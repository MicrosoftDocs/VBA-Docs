---
title: PageSetup.LayoutMode property (Word)
keywords: vbawd10.chm158400643
f1_keywords:
- vbawd10.chm158400643
ms.prod: word
api_name:
- Word.PageSetup.LayoutMode
ms.assetid: 9b5eb10a-0d90-5330-8738-f70efbae39fe
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.LayoutMode property (Word)

Returns or sets the layout mode for the current document. Read/write  **WdLayoutMode**.


## Syntax

_expression_. `LayoutMode`

_expression_ Required. A variable that represents a **[PageSetup](Word.PageSetup.md)** object.


## Example

This example sets the layout mode for the active document so that Microsoft Word automatically aligns typed text to a grid.


```vb
ActiveDocument.PageSetup.LayoutMode = wdLayoutModeGenko
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]