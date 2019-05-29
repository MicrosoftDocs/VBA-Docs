---
title: Options.AutoFormatAsYouTypeReplaceSymbols property (Word)
keywords: vbawd10.chm162988297
f1_keywords:
- vbawd10.chm162988297
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeReplaceSymbols
ms.assetid: 06d104d2-d8fa-8ef5-ba94-12b48f650c2a
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatAsYouTypeReplaceSymbols property (Word)

 **True** if two consecutive hyphens (--) are replaced with an en dash (-) or an em dash (—) as you type. Read/write **Boolean**.If the hyphens are typed with leading and trailing spaces, Word replaces the hyphens with an en dash; if there are no trailing spaces, the hyphens are replaced with an em dash.


## Syntax

_expression_. `AutoFormatAsYouTypeReplaceSymbols`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example turns on the replacement of hyphens with symbols as you type.


```vb
Options.AutoFormatAsYouTypeReplaceSymbols = True
```

This example returns the status of the Symbol characters (--) with symbols (—) option on the AutoFormat As You Type tab in the AutoCorrect dialog box (Tools menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeReplaceSymbols
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]