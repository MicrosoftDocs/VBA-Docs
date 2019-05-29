---
title: Options.MonthNames property (Word)
keywords: vbawd10.chm162988434
f1_keywords:
- vbawd10.chm162988434
ms.prod: word
api_name:
- Word.Options.MonthNames
ms.assetid: 265bee60-26ac-a6f5-4950-494ce6eff215
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MonthNames property (Word)

Returns or sets the direction for conversion between Hangul and Hanja. Read/write  **WdMonthNames**.


## Syntax

_expression_. `MonthNames`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to convert from Hangul to Hanja by default.


```vb
Options.MultipleWordConversionsMode = wdHangulToHanja
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]