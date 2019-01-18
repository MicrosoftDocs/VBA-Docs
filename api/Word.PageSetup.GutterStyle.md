---
title: PageSetup.GutterStyle property (Word)
keywords: vbawd10.chm158400641
f1_keywords:
- vbawd10.chm158400641
ms.prod: word
api_name:
- Word.PageSetup.GutterStyle
ms.assetid: cf2dafc3-1f08-d60d-830b-80ee921ee4cd
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.GutterStyle property (Word)

Returns or sets whether Microsoft Word uses gutters for the current document based on a right-to-left language or a left-to-right language. Read/write  **WdGutterStyleOld**.


## Syntax

 _expression_. `GutterStyle`

 _expression_ Required. A variable that represents a '[PageSetup](Word.PageSetup.md)' object.


## Example

This example sets the current document to follow a gutter style for a right-to-left language document.


```vb
ActiveDocument.PageSetup.GutterStyle = wdGutterStyleBidi
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]