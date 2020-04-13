---
title: EmailOptions.MarkComments property (Word)
keywords: vbawd10.chm165347435
f1_keywords:
- vbawd10.chm165347435
ms.prod: word
api_name:
- Word.EmailOptions.MarkComments
ms.assetid: 792e77b2-ba00-2b2b-c81b-7d00dad702cd
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.MarkComments property (Word)

 **True** if Microsoft Word marks the user's comments in email messages. Read/write **Boolean**.


## Syntax

_expression_. `MarkComments`

 _expression_ An expression that returns an '[EmailOptions](Word.EmailOptions.md)' object.


## Remarks

This property marks comments with the value of the **[MarkCommentsWith](Word.EmailOptions.MarkCommentsWith.md)** property. The default value of the **MarkCommentsWith** property is the value of the **[UserName](Word.Application.UserName.md)** property.


## Example

This example sets Word to mark comments in email messages with the initials "WK."


```vb
Application.EmailOptions.MarkCommentsWith = "WK" 
Application.EmailOptions.MarkComments = True
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]