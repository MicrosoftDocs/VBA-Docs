---
title: ParagraphFormat.HangingPunctuation property (Word)
keywords: vbawd10.chm156434551
f1_keywords:
- vbawd10.chm156434551
ms.prod: word
api_name:
- Word.ParagraphFormat.HangingPunctuation
ms.assetid: 9dc481f6-65fd-35f3-0765-087996aa6564
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.HangingPunctuation property (Word)

 **True** if hanging punctuation is enabled for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long**.


## Syntax

 _expression_. `HangingPunctuation`

 _expression_ A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Example

This example enables hanging punctuation for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).HangingPunctuation = True
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]