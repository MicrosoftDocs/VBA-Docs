---
title: Options.AddControlCharacters property (Word)
keywords: vbawd10.chm162988439
f1_keywords:
- vbawd10.chm162988439
ms.prod: word
api_name:
- Word.Options.AddControlCharacters
ms.assetid: 42d2e513-86a1-e8e3-8bc3-c133d90c3d2a
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AddControlCharacters property (Word)

 **True** if Microsoft Word adds bidirectional control characters when cutting and copying text. Read/write **Boolean**.


## Syntax

 _expression_. `AddControlCharacters`

 _expression_ An expression that returns an '[Options](Word.Options.md)' object.


## Example

This example sets Word to add bidirectional control characters when cutting and copying text.


```vb
Options.AddControlCharacters = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]