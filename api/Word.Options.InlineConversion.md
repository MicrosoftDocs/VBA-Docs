---
title: Options.InlineConversion property (Word)
keywords: vbawd10.chm162988118
f1_keywords:
- vbawd10.chm162988118
ms.prod: word
api_name:
- Word.Options.InlineConversion
ms.assetid: ee8d7237-86b0-74bd-ed19-dd09e29665d8
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.InlineConversion property (Word)

 **True** if Microsoft Word displays an unconfirmed character string in the Japanese Input Method Editor (IME) as an insertion between existing (confirmed) character strings. Read/write **Boolean**.


## Syntax

 _expression_. `InlineConversion`

 _expression_ An expression that returns an '[Options](Word.Options.md)' object.


## Example

This example sets Microsoft Word to display an unconfirmed character string in the Japanese Input Method Editor (IME) as an insertion between existing (confirmed) character strings.


```vb
Options.InlineConversion = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]