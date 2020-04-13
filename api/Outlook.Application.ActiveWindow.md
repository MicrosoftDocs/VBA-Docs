---
title: Application.ActiveWindow method (Outlook)
keywords: vbaol11.chm726
f1_keywords:
- vbaol11.chm726
ms.prod: outlook
api_name:
- Outlook.Application.ActiveWindow
ms.assetid: 5f5b4e8b-61e4-417b-6b0c-14d1ccb41594
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActiveWindow method (Outlook)

Returns an object representing the current Microsoft Outlook window on the desktop, either an **[Explorer](Outlook.Explorer.md)** or an **[Inspector](Outlook.Inspector.md)** object.


## Syntax

_expression_.**ActiveWindow**

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Return value

An **Object** that represents the current Outlook window on the desktop. Returns **Nothing** if no Outlook explorer or inspector is open.


## Example

This Microsoft Visual Basic for Applications (VBA) example minimizes the topmost Outlook window if it is an inspector window.


```vb
Sub MinimizeActiveWindow() 
 
 If TypeName(Application.ActiveWindow) = "Inspector" Then 
 
 Application.ActiveWindow.WindowState = olMinimized 
 
 End If 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
