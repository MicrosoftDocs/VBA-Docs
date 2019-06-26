---
title: Project.Activate method (Project)
ms.prod: project-server
api_name:
- Project.Project.Activate
ms.assetid: 965ad204-9f56-591f-91a1-7c42ded480cf
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Activate method (Project)

Activates the project.


## Syntax

_expression_.**Activate**

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following examples activate the next and previous projects, respectively.


```vb
Sub ProjectNext() 
    If ActiveProject.Index < Projects.Count Then 
        Projects(ActiveProject.Index + 1).Activate 
    Else 
        Projects(1).Activate 
    End If 
End Sub 
 
Sub ProjectPrevious() 
    If ActiveProject.Index > 1 Then 
        Projects(ActiveProject.Index - 1).Activate 
    Else 
         Projects(Projects.Count).Activate 
    End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]