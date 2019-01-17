---
title: Paragraph.LineUnitAfter property (Word)
keywords: vbawd10.chm156696706
f1_keywords:
- vbawd10.chm156696706
ms.prod: word
api_name:
- Word.Paragraph.LineUnitAfter
ms.assetid: 08abe0e4-4171-9d00-aedc-f714e4f2e60d
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.LineUnitAfter property (Word)

Returns or sets the amount of spacing (in gridlines) after the specified paragraph. Read/write  **Single**.


## Syntax

 _expression_. `LineUnitAfter`

 _expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example sets the spacing after the first paragraph in the active document to one gridline.


```vb
ActiveDocument.Paragraphs(1).LineUnitAfter = 1
```


## See also


[Paragraph Object](Word.Paragraph.md)

