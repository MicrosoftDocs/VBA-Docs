---
title: CoAuthoring.Locks property (Word)
keywords: vbawd10.chm254869509
f1_keywords:
- vbawd10.chm254869509
ms.prod: word
api_name:
- Word.CoAuthoring.Locks
ms.assetid: cf8feb0f-3617-c239-08de-ac6f8fc71b6e
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthoring.Locks property (Word)

Returns a  **[CoAuthLocks](Word.CoAuthLocks.md)** collection that represents the locks in the document. Read-only.


## Syntax

_expression_. `Locks`

 _expression_ An expression that returns a '[CoAuthoring](Word.CoAuthoring.md)' object.


## Example

The following code example displays the number of locks in the active document.


```vb
MsgBox "There are " & _ 
    ActiveDocument.CoAuthoring.Locks.Count & _ 
    " locks in the active document."
```


## See also


[CoAuthoring Object](Word.CoAuthoring.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]