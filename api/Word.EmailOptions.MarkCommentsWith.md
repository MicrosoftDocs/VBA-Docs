---
title: EmailOptions.MarkCommentsWith property (Word)
keywords: vbawd10.chm165347434
f1_keywords:
- vbawd10.chm165347434
ms.prod: word
api_name:
- Word.EmailOptions.MarkCommentsWith
ms.assetid: f10ce322-5ac5-f431-80c9-5c00a0892e2e
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.MarkCommentsWith property (Word)

Returns or sets the string with which Microsoft Word marks comments in email messages. Read/write  **String**.


## Syntax

_expression_. `MarkCommentsWith`

 _expression_ An expression that returns an '[EmailOptions](Word.EmailOptions.md)' object.


## Remarks

The default value is the value of the  **[UserName](Word.Application.UserName.md)** property.


## Example

This example sets Word to mark comments in email messages with the initials "WK."


```vb
Application.EmailOptions.MarkCommentsWith = "WK" 
Application.EmailOptions.MarkComments = True
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]