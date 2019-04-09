---
title: EmailOptions.AutoFormatAsYouTypeDeleteAutoSpaces property (Word)
keywords: vbawd10.chm165347630
f1_keywords:
- vbawd10.chm165347630
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeDeleteAutoSpaces
ms.assetid: d04465fa-2a63-7cb8-1163-868e454d832b
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.AutoFormatAsYouTypeDeleteAutoSpaces property (Word)

 **True** for Microsoft Word to automatically delete spaces inserted between Japanese and Latin text as you type. Read/write.


## Syntax

_expression_. `AutoFormatAsYouTypeDeleteAutoSpaces`

_expression_ Required. A variable that represents an '[EmailOptions](Word.EmailOptions.md)' collection.


## Example

This example sets Microsoft Word to automatically delete spaces inserted between Japanese and Latin text as you type.


```vb
Sub AutoDeleteSpaces() 
 Options.AutoFormatAsYouTypeDeleteAutoSpaces = True 
End Sub
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]