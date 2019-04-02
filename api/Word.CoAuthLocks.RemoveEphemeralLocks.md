---
title: CoAuthLocks.RemoveEphemeralLocks method (Word)
keywords: vbawd10.chm180486147
f1_keywords:
- vbawd10.chm180486147
ms.prod: word
api_name:
- Word.CoAuthLocks.RemoveEphemeralLocks
ms.assetid: fc894f97-b84c-8410-1847-ef2c3ad97300
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthLocks.RemoveEphemeralLocks method (Word)

Removes ephemeral locks from the document.


## Syntax

_expression_. `RemoveEphemeralLocks`

 _expression_ An expression that returns a '[CoAuthLocks](Word.CoAuthLocks.md)' object.


## Return value

Nothing


## Remarks

Ephemeral locks are automatically applied to a range when a co author begins to edit a document that has co authoring enabled.


## Example

The following code example removes all ephemeral locks from the active document.


```vb
ActiveDocument.CoAuthoring.Locks.RemoveEphemeralLocks  

```


## See also


[CoAuthLocks Object](Word.CoAuthLocks.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]