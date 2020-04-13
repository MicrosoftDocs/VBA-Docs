---
title: Task.OutlineChildren property (Project)
ms.prod: project-server
api_name:
- Project.Task.OutlineChildren
ms.assetid: e5e6f306-a0ea-d7b0-b627-3e8384705d62
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.OutlineChildren property (Project)

Gets a **[Tasks](Project.Task.md)** collection representing the children of a task in the outline structure. Read-only **Tasks**.


## Syntax

_expression_. `OutlineChildren`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


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