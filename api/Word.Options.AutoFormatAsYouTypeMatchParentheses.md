---
title: Options.AutoFormatAsYouTypeMatchParentheses property (Word)
keywords: vbawd10.chm162988332
f1_keywords:
- vbawd10.chm162988332
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeMatchParentheses
ms.assetid: f5f816db-8123-df7c-54cc-3e8ec6550207
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatAsYouTypeMatchParentheses property (Word)

 **True** for Microsoft Word to automatically correct improperly paired parentheses. Read/write.


## Syntax

_expression_. `AutoFormatAsYouTypeMatchParentheses`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to automatically correct improperly paired parentheses as you type.


```vb
Sub AutoMatchParentheses() 
 Options.AutoFormatAsYouTypeMatchParentheses = True 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]