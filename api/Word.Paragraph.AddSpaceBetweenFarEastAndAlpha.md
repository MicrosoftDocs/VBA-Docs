---
title: Paragraph.AddSpaceBetweenFarEastAndAlpha property (Word)
keywords: vbawd10.chm156696697
f1_keywords:
- vbawd10.chm156696697
ms.prod: word
api_name:
- Word.Paragraph.AddSpaceBetweenFarEastAndAlpha
ms.assetid: 3bcf9e22-42d1-0dbf-bbff-eb024db420e4
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.AddSpaceBetweenFarEastAndAlpha property (Word)

 **True** if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long**.


## Syntax

_expression_. `AddSpaceBetweenFarEastAndAlpha`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example sets Microsoft Word to automatically add spaces between Japanese and Latin text for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).AddSpaceBetweenFarEastAndAlpha = True
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]