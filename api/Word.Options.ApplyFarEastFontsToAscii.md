---
title: Options.ApplyFarEastFontsToAscii property (Word)
keywords: vbawd10.chm162988359
f1_keywords:
- vbawd10.chm162988359
ms.prod: word
api_name:
- Word.Options.ApplyFarEastFontsToAscii
ms.assetid: b0487311-42ad-f87a-8f72-da47d37f71d0
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.ApplyFarEastFontsToAscii property (Word)

 **True** if Microsoft Word applies East Asian fonts to Latin text. Read/write **Boolean**.


## Syntax

_expression_. `ApplyFarEastFontsToAscii`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

This property applies only when you have selected an East Asian language for editing. If this property is  **False** and you apply an East Asian font to a specified range, Word will not apply the font to any Latin text in the range.


## Example

This example sets Microsoft Word to apply East Asian fonts to Latin text.


```vb
Options.ApplyFarEastFontsToAscii = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]