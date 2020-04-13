---
title: ProtectedViewWindows object (Word)
ms.prod: word
api_name:
- Word.ProtectedViewWindows
ms.assetid: 62c2f4d5-1080-548e-730b-388308144dfe
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindows object (Word)

A collection of all the [ProtectedViewWindow](Word.ProtectedViewWindow.md) objects that are currently open in Word.


## Remarks

Use the **ProtectedViewWindows** property to return the **ProtectedViewWindows** collection.


## Example

The following code example displays the number of Protected View windows that are open.


```vb
MsgBox "There are " & ProtectedViewWindows.Count & _ 
 " Protected View windows currently open."
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]