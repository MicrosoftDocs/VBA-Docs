---
title: Selection.NoProofing property (Word)
keywords: vbawd10.chm158663661
f1_keywords:
- vbawd10.chm158663661
ms.prod: word
api_name:
- Word.Selection.NoProofing
ms.assetid: 5feca11c-5afa-80aa-b854-bab86b49a749
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.NoProofing property (Word)

 **True** if the spelling and grammar checker ignores the specified text. Returns **wdUndefined** if the **NoProofing** property is set to **True** for only some of the specified text. Read/write **Long**.


## Syntax

_expression_. `NoProofing`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example marks the current selection to be ignored by the spelling and grammar checker.


```vb
Selection.NoProofing = True
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]