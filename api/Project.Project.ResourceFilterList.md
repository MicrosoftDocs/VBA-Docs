---
title: Project.ResourceFilterList property (Project)
ms.prod: project-server
api_name:
- Project.Project.ResourceFilterList
ms.assetid: d515691a-2f8c-ed61-4844-3a938c658847
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.ResourceFilterList property (Project)

Gets a **[List](Project.List.md)** object representing all resource filters in the project. Read-only **List**.


## Syntax

_expression_. `ResourceFilterList`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following example lists all the resource filters in the active project.


```vb
Sub SeeAllResFilters() 
 
 Dim Temp As Variant 
 Dim ResFilterNames As String 
 
 For Each Temp In ActiveProject.ResourceFilterList 
 ResFilterNames = ResFilterNames & vbCrLf & Temp 
 Next Temp 
 
 MsgBox ResFilterNames 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]