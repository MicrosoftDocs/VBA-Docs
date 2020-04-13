---
title: ProtectedViewWindow.Height property (Word)
keywords: vbawd10.chm231735301
f1_keywords:
- vbawd10.chm231735301
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Height
ms.assetid: c3b423c9-25d4-3fc9-06b5-a7f8b88650d7
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.Height property (Word)

Returns or sets the height of the Protected View window. Read/write  **Long**.


## Syntax

_expression_.**Height**

 _expression_ An expression that returns a '[ProtectedViewWindow](Word.ProtectedViewWindow.md)' object.


## Remarks

You cannot set this property if the window is maximized or minimized. Use the **UsableHeight** property of the Application object to determine the maximum size for the window. Use the WindowState property to determine the window state.


## Example

The following code example changes the height of the active Protected View window to fill the application window area.


```vb
With ActiveProtectedViewWindow 
 .WindowState = wdWindowStateNormal 
 .Height = Application.UsableHeight 
End With
```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]