---
title: Options.AutoFormatAsYouTypeReplacePlainTextEmphasis property (Word)
keywords: vbawd10.chm162988300
f1_keywords:
- vbawd10.chm162988300
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeReplacePlainTextEmphasis
ms.assetid: 7c01c19a-1c3e-6bea-1979-ebd524bdf981
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatAsYouTypeReplacePlainTextEmphasis property (Word)

 **True** if manual emphasis characters are automatically replaced with character formatting as you type. For example, "*bold*" is changed to " **bold** " and "_underline_" is changed to "underline." Read/write **Boolean**.


## Syntax

_expression_. `AutoFormatAsYouTypeReplacePlainTextEmphasis`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example turns on the replacement of manual emphasis characters with character formatting.


```vb
Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = True
```

This example returns the status of the *Bold* and _underline_ with real formatting option on the AutoFormat As You Type tab in the AutoCorrect dialog box (Tools menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = _ 
 Options.AutoFormatAsYouTypeReplacePlainTextEmphasis
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]