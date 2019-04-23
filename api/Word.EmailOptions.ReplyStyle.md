---
title: EmailOptions.ReplyStyle property (Word)
keywords: vbawd10.chm165347438
f1_keywords:
- vbawd10.chm165347438
ms.prod: word
api_name:
- Word.EmailOptions.ReplyStyle
ms.assetid: adb778ca-8943-4f30-48d8-98336ea81ea7
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.ReplyStyle property (Word)

Returns a  **[Style](Word.Style.md)** object that represents the style used when replying to email messages.


## Syntax

_expression_. `ReplyStyle`

 _expression_ An expression that returns an '[EmailOptions](Word.EmailOptions.md)' object.


## Example

This example displays the name of the default style used when replying to email messages.


```vb
MsgBox Application.EmailOptions.ReplyStyle.NameLocal
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]