---
title: Window.WindowState property (Project)
ms.prod: project-server
api_name:
- Project.Window.WindowState
ms.assetid: b1c0616c-7377-356e-446d-ee2d2f490e15
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.WindowState property (Project)

Gets or sets the state the window, where the state is maximized or normal. Read/write  **PjWindowState**.


## Syntax

_expression_.**WindowState**

_expression_ A variable that represents a [Window](./Project.Window.md) object.


## Remarks

The **WindowState** property can be one of the following **[PjWindowState](Project.PjWindowState.md)** constants: **pjMaximized** or **pjNormal**. The **pjMinimized** value has no effect on a window within the Project application.

To change the state of the application window, use the **[WindowState](Project.Application.WindowState.md)** property of the **Application** object.


## Example

The following example maximizes the active window.


```vb
Sub MaximizeProjectWindow() 
 ActiveWindow.WindowState = pjMaximized 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]