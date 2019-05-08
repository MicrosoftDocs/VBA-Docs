---
title: ProtectedViewWindow.Activate method (Word)
keywords: vbawd10.chm231735396
f1_keywords:
- vbawd10.chm231735396
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Activate
ms.assetid: a784fceb-38b9-2fc4-6c71-fcfb17b53dfe
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.Activate method (Word)

Activates the specified Protected View window.


## Syntax

_expression_.**Activate**

 _expression_ An expression that returns a '[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)' object.


## Return value

Nothing


## Example

The following code example activates the next Protected View window in the [ProtectedViewWindows](Word.ProtectedViewWindows.md) collection.


```vb
Dim pvWindow As ProtectedViewWindow 
 
' At least one document must be open in protected view for this statement to execute. 
ProtectedViewWindows(1).Activate
```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]