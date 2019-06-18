---
title: Window.Move method (Publisher)
keywords: vbapb10.chm262163
f1_keywords:
- vbapb10.chm262163
ms.prod: publisher
api_name:
- Publisher.Window.Move
ms.assetid: a33b213b-6549-abf7-0217-041b469b798a
ms.date: 06/18/2019
localization_priority: Normal
---


# Window.Move method (Publisher)

Moves the active document window.


## Syntax

_expression_.**Move** (_Left_, _Top_)

_expression_ A variable that represents a **[Window](Publisher.Window.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Left_|Required| **Long**|The horizontal screen position of the specified window.|
|_Top_|Required| **Long**|The vertical screen position of the specified window.|

## Remarks

If the application window is either maximized or minimized, this method returns an error.


## Example

This example checks the state of the application window, and if it is neither maximized nor minimized, moves the window to the upper-left corner of the screen.

```vb
Sub MoveWindow() 
 With ActiveWindow 
 If .WindowState = pbWindowStateNormal Then 
 .Move Left:=50, Top:=50 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]