---
title: Conflict.Reject method (Word)
keywords: vbawd10.chm78708838
f1_keywords:
- vbawd10.chm78708838
ms.prod: word
api_name:
- Word.Conflict.Reject
ms.assetid: 9bd4fa93-4bae-e2a8-ef6e-b3116542cad4
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflict.Reject method (Word)

Rejects the user change, removes the conflict, and accepts the server copy of the change for the conflict.


## Syntax

_expression_. `Reject`

 _expression_ An expression that returns a [Conflict](./Word.Conflict.md) object.


## Return value

Nothing


## Remarks

The  **Reject** method rejects the user version of a conflict and accepts the version that is currently on the server.


## Example

The following code example rejects all the conflicts in the active document.


```vb
Dim conf As Conflict 
 
For Each conf In ActiveDocument.CoAuthoring.Conflicts 
 conf.Reject 
Next conf
```

Alternatively, you can use the [RejectAll](Word.Conflicts.RejectAll.md) method of the [Conflicts](Word.Conflicts.md) collection object to reject all the conflicts in a document, as shown in the following code example.




```vb
ActiveDocument.CoAuthoring.Conflicts.RejectAll
```


## See also


[Conflict Object](Word.Conflict.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]