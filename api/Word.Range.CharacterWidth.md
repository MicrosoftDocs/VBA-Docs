---
title: Range.CharacterWidth property (Word)
keywords: vbawd10.chm157155654
f1_keywords:
- vbawd10.chm157155654
ms.prod: word
api_name:
- Word.Range.CharacterWidth
ms.assetid: 83eadb2b-5c79-d246-d1f1-fd6a9e1f4bd8
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.CharacterWidth property (Word)

Returns or sets the character width of the specified range. Read/write  **WdCharacterWidth**.


## Syntax

_expression_. `CharacterWidth`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Example

This example converts the current selection to half-width characters.


```vb
Selection.Range.CharacterWidth = wdWidthHalfWidth
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]