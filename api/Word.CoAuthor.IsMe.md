---
title: CoAuthor.IsMe property (Word)
keywords: vbawd10.chm81068035
f1_keywords:
- vbawd10.chm81068035
ms.prod: word
api_name:
- Word.CoAuthor.IsMe
ms.assetid: bf6b8282-e114-8b6f-9e89-3bd93662d84e
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthor.IsMe property (Word)

Returns true if this author represents the current user. Read-only. 


## Syntax

_expression_. `IsMe`

 _expression_ An expression that returns a [CoAuthor](./Word.CoAuthor.md) object.


## Example

The following code example checks the active document to see if the first co author in the CoAuthors collection is the current user.


```vb
If ActiveDocument.CoAuthoring.Authors(1).IsMe Then 
MsgBox "The current user is the first coauthor." 
End If
```


## See also


[CoAuthor Object](Word.CoAuthor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]