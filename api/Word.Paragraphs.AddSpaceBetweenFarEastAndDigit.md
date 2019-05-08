---
title: Paragraphs.AddSpaceBetweenFarEastAndDigit property (Word)
keywords: vbawd10.chm156762234
f1_keywords:
- vbawd10.chm156762234
ms.prod: word
api_name:
- Word.Paragraphs.AddSpaceBetweenFarEastAndDigit
ms.assetid: 7ecf08cb-1d5b-f833-de21-7b1c00cc3937
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.AddSpaceBetweenFarEastAndDigit property (Word)

 **True** if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long**.


## Syntax

_expression_. `AddSpaceBetweenFarEastAndDigit`

_expression_ A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example sets Microsoft Word to automatically add spaces between Japanese text and numbers for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).AddSpaceBetweenFarEastAndDigit = True
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]