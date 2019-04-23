---
title: Paragraphs.OutlineDemoteToBody method (Word)
keywords: vbawd10.chm156762438
f1_keywords:
- vbawd10.chm156762438
ms.prod: word
api_name:
- Word.Paragraphs.OutlineDemoteToBody
ms.assetid: 26eedf4b-fcca-d065-40c2-76e191608678
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.OutlineDemoteToBody method (Word)

Demotes the specified paragraph or paragraphs to body text by applying the Normal style.


## Syntax

_expression_. `OutlineDemoteToBody`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example demotes the selected paragraphs to body text by applying the Normal style.


```vb
Selection.Paragraphs.OutlineDemoteToBody
```

This example switches the active window to outline view and demotes all selected paragraphs to body text.




```vb
ActiveDocument.ActiveWindow.View.Type = wdOutlineView 
Selection.Paragraphs.OutlineDemoteToBody
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]