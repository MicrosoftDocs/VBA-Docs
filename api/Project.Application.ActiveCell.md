---
title: Application.ActiveCell property (Project)
keywords: vbapj.chm131368
f1_keywords:
- vbapj.chm131368
ms.prod: project-server
api_name:
- Project.Application.ActiveCell
ms.assetid: 880931d8-fc23-7938-e4fe-bd800eeae318
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActiveCell property (Project)

Gets a  **[Cell](Project.Cell.md)** object that represents the active cell. Read-only **Cell**.


## Syntax

_expression_.**ActiveCell**

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Example

The following example displays the names of the resources assigned to the selected task. The example assumes a task view is the active view and the active cell is in a task row.


```vb
Sub ResourceNames() 
 
 Dim A As Assignment 
 
 For Each A In ActiveCell.Task.Assignments 
 MsgBox A.ResourceName 
 Next A 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]