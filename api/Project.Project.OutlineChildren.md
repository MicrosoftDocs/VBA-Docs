---
title: Project.OutlineChildren property (Project)
ms.prod: project-server
api_name:
- Project.Project.OutlineChildren
ms.assetid: f0feaf89-04ad-4523-7b15-eff6573f6ddd
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.OutlineChildren property (Project)

Gets a **[Tasks](Project.Task.md)** collection representing the children of a task in the outline structure. Read-only **Tasks**.


## Syntax

_expression_. `OutlineChildren`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following example displays the names of all tasks at the same outline level as the selected task.


```vb
Sub Siblings() 
 
 Dim MyParent As Task 
 Dim Sibling As Task 
 Dim Temp As String 
 
 Set MyParent = ActiveCell.Task.OutlineParent 
 
 For Each Sibling In MyParent.OutlineChildren 
 Temp = Sibling.Name & ListSeparator & " " & Temp 
 Next Sibling 
 
 Temp = Left$(Temp, Len(Temp) - Len(ListSeparator & " ")) 
 MsgBox Temp 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]