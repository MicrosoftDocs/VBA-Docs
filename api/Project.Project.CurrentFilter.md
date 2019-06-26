---
title: Project.CurrentFilter property (Project)
keywords: vbapj.chm131700
f1_keywords:
- vbapj.chm131700
ms.prod: project-server
api_name:
- Project.Project.CurrentFilter
ms.assetid: b97e43ac-2167-80f0-bf5e-609a08f42fd9
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.CurrentFilter property (Project)

Gets the name of the active filter for a project. Read-only  **String**.


## Syntax

_expression_. `CurrentFilter`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following example displays the names of the active view, table, and filter in a dialog box


```vb
Sub ViewDetails()

    Dim Temp As String

    Temp = "View: " & ActiveProject.CurrentView & vbCrLf 
    Temp = Temp & "Table:" & ActiveProject.CurrentTable & vbCrLf 
    Temp = Temp & "Filter: " & ActiveProject.CurrentFilter 
    MsgBox Temp 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]