---
title: CoAuthLock.Unlock method (Word)
keywords: vbawd10.chm260046854
f1_keywords:
- vbawd10.chm260046854
ms.prod: word
api_name:
- Word.CoAuthLock.Unlock
ms.assetid: 7f64431b-8b1b-60b3-142c-5251dce1a8a1
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthLock.Unlock method (Word)

Unlocks the specified lock.


## Syntax

_expression_. `Unlock`

 _expression_ An expression that returns a '[CoAuthLock](Word.CoAuthLock.md)' object.


## Return value

Nothing


## Remarks

The **Unlock** method unlocks the specified lock even if it belongs to another user other than the current user.


## Example

The following code example unlocks all locks in the active document.


```vb
Dim myLock as CoAuthLock 
 
For Each myLock In ActiveDocument.CoAuthoring.Locks 
   myLock.Unlock     
Next myLock
```


## See also


[CoAuthLock Object](Word.CoAuthLock.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]