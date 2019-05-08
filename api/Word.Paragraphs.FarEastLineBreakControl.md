---
title: Paragraphs.FarEastLineBreakControl property (Word)
keywords: vbawd10.chm156762229
f1_keywords:
- vbawd10.chm156762229
ms.prod: word
api_name:
- Word.Paragraphs.FarEastLineBreakControl
ms.assetid: 4049497d-430b-8951-3d50-53a83e32c75d
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.FarEastLineBreakControl property (Word)

 **True** if Microsoft Word applies East Asian line-breaking rules to the specified paragraphs. Returns **wdUndefined** if the **FarEastLineBreakControl** property is set to **True** for only some of the specified paragraphs. Read/write **Long**.


## Syntax

_expression_. `FarEastLineBreakControl`

_expression_ A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example sets Word to apply East Asian line-breaking rules to the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).FarEastLineBreakControl = True
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]