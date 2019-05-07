---
title: Document.Email property (Word)
keywords: vbawd10.chm158007615
f1_keywords:
- vbawd10.chm158007615
ms.prod: word
api_name:
- Word.Document.Email
ms.assetid: dd4f6a41-3ee6-c1bf-3a2c-e00a342e0009
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Email property (Word)

Returns an  **[Email](Word.Email.md)** object that contains all the email-related properties of the current document. Read-only.


## Syntax

_expression_. `Email`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example returns the name of the style associated with the current email author.


```vb
MsgBox ActiveDocument.Email _ 
 .CurrentEmailAuthor.Style.NameLocal
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]