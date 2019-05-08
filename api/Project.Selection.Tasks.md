---
title: Selection.Tasks property (Project)
ms.prod: project-server
api_name:
- Project.Selection.Tasks
ms.assetid: 8f58ea8e-a3a1-f5aa-ad5d-6447fe777453
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Tasks property (Project)

Gets a  **[Tasks](Project.Task.md)** collection representing the tasks in the selection. Read-only **Tasks**.


## Syntax

_expression_. `Tasks`

_expression_ A variable that represents a [Selection](./Project.Selection.md) object.


## Example

The following example displays the name of every task in the selection.


```vb
Sub TaskNames() 
 
 Dim T As Task, Names As String 
 
 For Each T In ActiveSelection.Tasks 
 Names = Names & T.Name & vbCrLf 
 Next T 
 
 MsgBox Names 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]