---
title: ParagraphFormat.FirstLineIndent property (Word)
keywords: vbawd10.chm156434540
f1_keywords:
- vbawd10.chm156434540
ms.prod: word
api_name:
- Word.ParagraphFormat.FirstLineIndent
ms.assetid: a9a94019-537c-942d-c388-06b228fd5463
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.FirstLineIndent property (Word)

Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single**.


## Syntax

 _expression_. `FirstLineIndent`

 _expression_ A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Example

This example sets a first-line indent of 1 inch for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).FirstLineIndent = _ 
 InchesToPoints(1)
```

This example sets a hanging indent of 0.5 inch for the second paragraph in the active document. The InchesToPoints method is used to convert inches to points.




```vb
ActiveDocument.Paragraphs(2).FirstLineIndent = _ 
 InchesToPoints(-0.5)
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

