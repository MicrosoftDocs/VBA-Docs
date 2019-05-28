---
title: Options.InterpretHighAnsi property (Word)
keywords: vbawd10.chm162988450
f1_keywords:
- vbawd10.chm162988450
ms.prod: word
api_name:
- Word.Options.InterpretHighAnsi
ms.assetid: c093469b-c9ef-0b37-fc40-7b1ae17ce72e
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.InterpretHighAnsi property (Word)

Returns or sets the high-ANSI text interpretation behavior. Read/write  **WdHighAnsiText**.


## Syntax

_expression_. `InterpretHighAnsi`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Word to interpret all high-ANSI text as East Asian characters.


```vb
Options.InterpretHighAnsi = wdHighAnsiIsFarEast
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]