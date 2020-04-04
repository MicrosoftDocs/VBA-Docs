---
title: Task.OutlineParent property (Project)
ms.prod: project-server
api_name:
- Project.Task.OutlineParent
ms.assetid: 54dc7d2a-feb0-da23-5116-decf0f4388e9
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.OutlineParent property (Project)

Gets a **[Task](Project.Task.md)** object representing the parent of a task in the outline structure. Read-only **Task**.


## Syntax

_expression_. `OutlineParent`

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