---
title: Options.AutoFormatAsYouTypeReplaceQuotes property (Word)
keywords: vbawd10.chm162988296
f1_keywords:
- vbawd10.chm162988296
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeReplaceQuotes
ms.assetid: d0e2010c-efc3-f944-4daf-48f4ed36004b
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatAsYouTypeReplaceQuotes property (Word)

 **True** if straight quotation marks are automatically changed to smart (curly) quotation marks as you type. Read/write **Boolean**.


## Syntax

 _expression_. `AutoFormatAsYouTypeReplaceQuotes`

 _expression_ A variable that represents an '[Options](Word.Options.md)' collection.


## Example

This example turns on the automatic replacement of straight quotation marks with smart (curly) quotation marks as you type.


```vb
Options.AutoFormatAsYouTypeReplaceQuotes = True
```

This example returns the status of the Straight quotes with smart quotes option on the AutoFormat As You Type tab in the AutoCorrect dialog box (Tools menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatReplaceQuotes
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]