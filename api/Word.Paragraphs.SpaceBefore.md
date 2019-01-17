---
title: Paragraphs.SpaceBefore property (Word)
keywords: vbawd10.chm156762223
f1_keywords:
- vbawd10.chm156762223
ms.prod: word
api_name:
- Word.Paragraphs.SpaceBefore
ms.assetid: e526a660-96aa-acf3-2562-addb3e3af113
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.SpaceBefore property (Word)

Returns or sets the spacing (in points) before the specified paragraphs. Read/write  **Single**.


## Syntax

 _expression_. `SpaceBefore`

 _expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example sets the spacing before all paragraphs in the active document to 12 points.


```vb
ActiveDocument.Paragraphs.SpaceBefore = 12
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

