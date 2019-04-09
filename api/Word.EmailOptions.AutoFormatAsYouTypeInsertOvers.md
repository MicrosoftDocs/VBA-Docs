---
title: EmailOptions.AutoFormatAsYouTypeInsertOvers property (Word)
keywords: vbawd10.chm165347633
f1_keywords:
- vbawd10.chm165347633
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeInsertOvers
ms.assetid: 0c8b77a9-f6ed-1be5-bab8-dbab886812cd
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.AutoFormatAsYouTypeInsertOvers property (Word)

 **True** for Microsoft Word to automatically insert "以上" when the user enters "記" or "案". Read/write **Boolean**.


## Syntax

_expression_. `AutoFormatAsYouTypeInsertOvers`

_expression_ Required. A variable that represents an '[EmailOptions](Word.EmailOptions.md)' collection.


## Example

This example sets Microsoft Word to automatically insert "以上" when the user enters "記" or "案".


```vb
Options.AutoFormatAsYouTypeInsertOvers = True
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]