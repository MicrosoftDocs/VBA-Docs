---
title: Range.HighlightColorIndex property (Word)
keywords: vbawd10.chm157155629
f1_keywords:
- vbawd10.chm157155629
ms.prod: word
api_name:
- Word.Range.HighlightColorIndex
ms.assetid: ff6e0f1a-8b37-1bdd-8da6-ac492d399ad2
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.HighlightColorIndex property (Word)

Returns or sets the highlight color for the specified range. Read/write  **WdColorIndex**.


## Syntax

_expression_. `HighlightColorIndex`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Example

This example removes highlight formatting from the selection.


```vb
Selection.Range.HighlightColorIndex = wdNoHighlight
```

This example applies yellow highlighting to each bookmark in the active document.




```vb
For Each abookmark In ActiveDocument.Bookmarks 
 abookmark.Range.HighlightColorIndex = wdYellow 
Next abookmark
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
