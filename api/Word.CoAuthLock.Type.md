---
title: CoAuthLock.Type property (Word)
keywords: vbawd10.chm260046849
f1_keywords:
- vbawd10.chm260046849
ms.prod: word
api_name:
- Word.CoAuthLock.Type
ms.assetid: a88c38de-bea1-1766-cb33-c86eb30ef98e
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthLock.Type property (Word)

Returns a [WdLockType](Word.WdLockType.md) constant that specifies the lock type. Read-only.


## Syntax

_expression_.**Type**

 _expression_ An expression that returns a '[CoAuthLock](Word.CoAuthLock.md)' object.


## Example

The following code example removes all the reservation locks in the active document.


```vb
Dim myLock As CoAuthLock 
 
For Each myLock In ActiveDocument.CoAuthoring.Locks 
    If myLock.Type = wdLockReservation Then 
        myLock.Unlock 
    End If 
Next myLock
```


## See also


[CoAuthLock Object](Word.CoAuthLock.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]