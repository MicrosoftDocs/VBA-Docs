---
title: CurrentProject.IsTrusted property (Access)
keywords: vbaac10.chm12730
f1_keywords:
- vbaac10.chm12730
ms.prod: access
api_name:
- Access.CurrentProject.IsTrusted
ms.assetid: c3d8b6f8-c79f-79ab-d4e0-0454f97ac937
ms.date: 02/27/2019
localization_priority: Normal
---


# CurrentProject.IsTrusted property (Access)

Gets whether macros and Visual Basic for Applications (VBA) code have been enabled in the current project. Read-only **Boolean**.


## Syntax

_expression_.**IsTrusted**

_expression_ A variable that represents a **[CurrentProject](Access.CurrentProject.md)** object.


## Example

The following example shows how to use the **IsTrusted** property in a macro to determine whether the database has been opened with trust enabled. If trust has been enabled, the Visual Basic for Applications (VBA) subroutine **Init** is called. Otherwise, the user is notified that the database has been opened in disabled mode.

```vb
    If [currentproject].[istrusted] Then
        RunCode
            Function Name =Init()

    Else
        MessageBox
            Message The application is opened in disabled mode. Please enable the application for full functionality.
            Beep Yes
            Type None
            Title Disabled Mode Check

    End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]