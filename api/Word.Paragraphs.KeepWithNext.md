---
title: Paragraphs.KeepWithNext property (Word)
keywords: vbawd10.chm156762215
f1_keywords:
- vbawd10.chm156762215
ms.prod: word
api_name:
- Word.Paragraphs.KeepWithNext
ms.assetid: a0083251-893b-5323-7b4f-03df6ac32822
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.KeepWithNext property (Word)

 **True** if the specified paragraphs remain on the same page as the paragraphs that follow it when Microsoft Word repaginates the document. Read/write **Long**.


## Syntax

_expression_. `KeepWithNext`

_expression_ A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

This property can be  **True**, **False**, or **wdUndefined**.


## Example

This example sets all paragraphs in the current selection to be on the same page.


```vb
Selection.Paragraphs.KeepWithNext = True
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]