---
title: CoAuthLock.Owner property (Word)
keywords: vbawd10.chm260046850
f1_keywords:
- vbawd10.chm260046850
ms.prod: word
api_name:
- Word.Owner
ms.assetid: 55158805-f9fe-6cb0-c13a-30207b5f6f2d
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthLock.Owner property (Word)

Returns the **[CoAuthor](Word.CoAuthor.md)** that owns the specified lock. Read-only.


## Syntax

_expression_. `Owner`

 _expression_ An expression that returns a '[CoAuthLock](Word.CoAuthLock.md)' object.


## Example

The following code example displays the name of the owner of each lock in the active document.


```vb
Dim myLock As CoAuthLock 
 
For Each myLock In ActiveDocument.CoAuthoring.Locks 
    MsgBox "The owner of this lock is " & _ 
    myLock.Owner.Name & "." 
Next myLock
```


## See also


[CoAuthLock Object](Word.CoAuthLock.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]