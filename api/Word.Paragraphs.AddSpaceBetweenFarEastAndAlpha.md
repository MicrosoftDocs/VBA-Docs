---
title: Paragraphs.AddSpaceBetweenFarEastAndAlpha property (Word)
keywords: vbawd10.chm156762233
f1_keywords:
- vbawd10.chm156762233
ms.prod: word
api_name:
- Word.Paragraphs.AddSpaceBetweenFarEastAndAlpha
ms.assetid: f101d2fa-f999-b9fb-84c1-3f060fab7ed0
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.AddSpaceBetweenFarEastAndAlpha property (Word)

 **True** if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long**.


## Syntax

_expression_. `AddSpaceBetweenFarEastAndAlpha`

_expression_ A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example sets Microsoft Word to automatically add spaces between Japanese and Latin text for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).AddSpaceBetweenFarEastAndAlpha = True
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]