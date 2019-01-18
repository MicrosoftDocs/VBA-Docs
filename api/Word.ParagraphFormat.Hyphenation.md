---
title: ParagraphFormat.Hyphenation property (Word)
keywords: vbawd10.chm156434545
f1_keywords:
- vbawd10.chm156434545
ms.prod: word
api_name:
- Word.ParagraphFormat.Hyphenation
ms.assetid: 185d00c0-3f19-bc98-9790-823b49d289b1
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.Hyphenation property (Word)

 **True** if the specified paragraphs are included in automatic hyphenation. **False** if the specified paragraphs are to be excluded from automatic hyphenation. Read/write **Long**.


## Syntax

 _expression_. `Hyphenation`

 _expression_ A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Remarks

This property can be  **True** , **False** or **wdUndefined**.


## Example

This example turns off automatic hyphenation for all paragraphs in the active document that have the Normal style.


```vb
ActiveDocument.Styles("Normal").ParagraphFormat.Hyphenation = False
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]