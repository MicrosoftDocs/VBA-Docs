---
title: Application.CommandBars property (Project)
ms.prod: project-server
api_name:
- Project.Application.CommandBars
ms.assetid: 80f57057-9bb3-018b-0e45-fd1423368091
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CommandBars property (Project)

Gets a **CommandBars** collection that represents all the command bars in the application. Read-only **CommandBars**.


## Syntax

_expression_.**CommandBars**

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

For more information, see the **CommandBars** collection object in the Microsoft Office Visual Basic Reference.


## Example

The following example deletes all custom command bars that aren't visible.


```vb
Sub RemoveCommandBars() 
 Dim Bar As CommandBar 
 
 For Each Bar In Application.CommandBars 
 If Not Bar.BuiltIn And Not Bar.Visible Then Bar.Delete 
 Next 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]