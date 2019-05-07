---
title: Email.CurrentEmailAuthor property (Word)
keywords: vbawd10.chm165478505
f1_keywords:
- vbawd10.chm165478505
ms.prod: word
api_name:
- Word.Email.CurrentEmailAuthor
ms.assetid: a317b265-f712-2954-aeaf-3a17da38d380
ms.date: 06/08/2017
localization_priority: Normal
---


# Email.CurrentEmailAuthor property (Word)

Returns an  **[EmailAuthor](Word.EmailAuthor.md)** object that represents the author of the current email message. Read-only.


## Syntax

_expression_. `CurrentEmailAuthor`

_expression_ A variable that represents a '[Email](Word.Email.md)' object.


## Example

This example returns the name of the style associated with the current email author.


```vb
MsgBox ActiveDocument.Email _ 
 .CurrentEmailAuthor.Style.NameLocal
```


## See also


[Email Object](Word.Email.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]