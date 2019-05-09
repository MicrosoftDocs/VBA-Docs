---
title: ProtectedViewWindow.Left property (Word)
keywords: vbawd10.chm231735298
f1_keywords:
- vbawd10.chm231735298
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Left
ms.assetid: 55ca42b8-bed4-3b7e-fd0b-66dc2ea936c3
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.Left property (Word)

Returns or sets a  **Long**, in [points](../language/glossary/vbe-glossary.md#point), that represents the horizontal position of the specified Protected View window. Read/write.


## Syntax

_expression_.**Left**

 _expression_ An expression that returns a '[ProtectedViewWindow](Word.ProtectedViewWindow.md)' object.


## Example

The following code example sets the horizontal position of the active Protected View window to 100 point.


```vb
With ActiveProtectedViewWindow 
 .WindowState = wdWindowStateNormal 
 .Left = 100 
 .Top = 0 
End With
```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]