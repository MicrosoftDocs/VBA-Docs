---
title: CoAuthLocks.Count property (Word)
keywords: vbawd10.chm180486145
f1_keywords:
- vbawd10.chm180486145
ms.prod: word
api_name:
- Word.CoAuthLocks.Count
ms.assetid: a082d159-8fd9-1f8d-0987-7755f2aa4d5e
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthLocks.Count property (Word)

Returns a  **Long** that represents the number of locks in the **[CoAuthLocks](Word.CoAuthLocks.md)** collection. Read-only.


## Syntax

_expression_.**Count**

 _expression_ An expression that returns a [CoAuthLocks](./Word.CoAuthLocks.md) object.


## Example

The following code example displays the number of locks in the active document.


```vb
MsgBox "The active document contains " & _ 
    ActiveDocument.CoAuthoring.Locks.Count & " locks."
```


## See also


[CoAuthLocks Object](Word.CoAuthLocks.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]