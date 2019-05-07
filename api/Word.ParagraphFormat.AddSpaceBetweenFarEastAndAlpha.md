---
title: ParagraphFormat.AddSpaceBetweenFarEastAndAlpha property (Word)
keywords: vbawd10.chm156434553
f1_keywords:
- vbawd10.chm156434553
ms.prod: word
api_name:
- Word.ParagraphFormat.AddSpaceBetweenFarEastAndAlpha
ms.assetid: 3575dab1-4a59-b20e-46e2-971389a3ec95
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.AddSpaceBetweenFarEastAndAlpha property (Word)

 **True** if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long**.


## Syntax

_expression_. `AddSpaceBetweenFarEastAndAlpha`

_expression_ A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Example

This example sets Microsoft Word to automatically add spaces between Japanese and Latin text for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).AddSpaceBetweenFarEastAndAlpha = True
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]