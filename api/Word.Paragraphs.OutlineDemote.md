---
title: Paragraphs.OutlineDemote method (Word)
keywords: vbawd10.chm156762437
f1_keywords:
- vbawd10.chm156762437
ms.prod: word
api_name:
- Word.Paragraphs.OutlineDemote
ms.assetid: 24650317-73a4-67a3-d7f4-dfc25bd75d2a
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.OutlineDemote method (Word)

Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraphs.


## Syntax

 _expression_. `OutlineDemote`

 _expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

If a paragraph is formatted with the Heading 2 style, this method demotes the paragraph by changing the style to Heading 3.


## Example

This example demotes the selected paragraphs.


```vb
Selection.Paragraphs.OutlineDemote
```

This example demotes all paragraphs in the active document.




```vb
ActiveDocument.Paragraphs.OutlineDemote
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]