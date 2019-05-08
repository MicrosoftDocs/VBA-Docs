---
title: ProtectedViewWindow.Width property (Word)
keywords: vbawd10.chm231735300
f1_keywords:
- vbawd10.chm231735300
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Width
ms.assetid: 607ec503-2096-4b4a-fce5-9979bea6c847
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.Width property (Word)

Returns or sets the width, in [points](../language/glossary/vbe-glossary.md#point), of the specified Protected View window. Read/write  **Long**.


## Syntax

_expression_.**Width**

 _expression_ An expression that returns a '[ProtectedViewWindow](Word.ProtectedViewWindow.md)' object.


## Example

The following code example changes the state, height, and width of the active Protected View window.


```vb
With ActiveProtectedViewWindow 
 .WindowState = wdWindowStateNormal 
 .Height = 400 
 .Width = 500 
End With 

```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]