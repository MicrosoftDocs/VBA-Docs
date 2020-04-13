---
title: Task.PercentWorkComplete property (Project)
ms.prod: project-server
api_name:
- Project.Task.PercentWorkComplete
ms.assetid: f1b1dc5e-843c-ca0f-72f1-f8d7cdf6edab
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.PercentWorkComplete property (Project)

Gets or sets the percentage of work complete for a task. Read-only for summary tasks. Read/write  **Variant**.


## Syntax

_expression_. `PercentWorkComplete`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Example

The following example sets the **Marked** property to **True** for each task in the active project with a percentage of work complete that exceeds the percentage specified by the user.


```vb
Sub MarkTasks() 
 
 Dim T As Task ' Task object used in For Each loop 
 Dim Entry As String ' Percentage entered by user 
 
 ' Prompt user for a percentage. 
 Entry = InputBox$("Mark tasks that exceed what percentage of work complete? (0-100)") 
 
 If Not IsNumeric(Entry) Then 
 MsgBox ("Please enter a number only.") 
 Exit Sub 
 ElseIf Entry < 0 Or Entry > 100 Then 
 MsgBox ("You did not enter a percentage from 0 to 100.") 
 Exit Sub 
 End If 
 
 ' Mark tasks with percentage of work complete greater than user entry. 
 For Each T In ActiveProject.Tasks 
 If T.PercentWorkComplete > Val(Entry) Then 
 T.Marked = True 
 Else 
 T.Marked = False 
 End If 
 Next T 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]