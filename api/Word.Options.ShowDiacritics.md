---
title: Options.ShowDiacritics property (Word)
keywords: vbawd10.chm162988437
f1_keywords:
- vbawd10.chm162988437
ms.prod: word
api_name:
- Word.Options.ShowDiacritics
ms.assetid: b06b6d5e-1606-20c3-7efb-212503bc2790
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.ShowDiacritics property (Word)

 **True** if diacritics are visible in a right-to-left language document. Read/write **Boolean**.


## Syntax

_expression_. `ShowDiacritics`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example hides diacritics in the current document.


```vb
Options.ShowDiacritics = False
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]