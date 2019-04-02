---
title: CoAuthoring.Me property (Word)
keywords: vbawd10.chm254869506
f1_keywords:
- vbawd10.chm254869506
ms.prod: word
api_name:
- Word.CoAuthoring.Me
ms.assetid: 19c2875f-07ba-15c3-a622-254344c6480f
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthoring.Me property (Word)

Returns a  **[CoAuthor](Word.CoAuthor.md)** object that represents the current user. Read-only.


## Syntax

_expression_. `Me`

 _expression_ An expression that returns a '[CoAuthoring](Word.CoAuthoring.md)' object.


## Example

The following code example gets the number of locks in the active document that are associated with the current user.


```vb
Dim coAuth As CoAuthor 
 
Set coAuth = ActiveDocument.CoAuthoring.Me 
MsgBox "The current user has " & coAuth.Locks.Count & _ 
" locks in the active document."
```


## See also


[CoAuthoring Object](Word.CoAuthoring.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]