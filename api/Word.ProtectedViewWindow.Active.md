---
title: ProtectedViewWindow.Active property (Word)
keywords: vbawd10.chm231735303
f1_keywords:
- vbawd10.chm231735303
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Active
ms.assetid: 8c301a06-aaca-4ecf-cf08-563b45810028
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.Active property (Word)

 **True** if the specified Protected View window is active. Read-only **Boolean**.


## Syntax

_expression_.**Active**

 _expression_ An expression that returns a [ProtectedViewWindow](./Word.ProtectedViewWindow.md) object.


## Example

The following code example activates the first Protected View window in the [ProtectedViewWindows](Word.ProtectedViewWindows.md) collection if the window is not currently active.


```vb
ProtectedViewWindows.Open FileName:="C:\MyFiles\MyDoc.doc" 

```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]