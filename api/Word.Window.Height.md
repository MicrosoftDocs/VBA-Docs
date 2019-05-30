---
title: Window.Height property (Word)
keywords: vbawd10.chm157417480
f1_keywords:
- vbawd10.chm157417480
ms.prod: word
api_name:
- Word.Window.Height
ms.assetid: 9b96ac83-57cc-4cb2-768b-2b5012c49bbc
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Height property (Word)

Returns or sets the height of the window (in points). Read/write **Long**.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[Window](Word.Window.md)** object.


## Remarks

You cannot set this property if the window is maximized or minimized. Use the  **UsableHeight** property of the **Application** object to determine the maximum size for the window. Use the **WindowState** property to determine the window state.


## Example

This example changes the height of the active window to fill the application window area.


```vb
With ActiveDocument.ActiveWindow 
 .WindowState = wdWindowStateNormal 
 .Height = Application.UsableHeight 
End With
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]