---
title: Period.Count property (Project)
ms.prod: project-server
api_name:
- Project.Period.Count
ms.assetid: 8b1caae6-2ae1-12c4-1f94-b52dcececd45
ms.date: 06/08/2017
localization_priority: Normal
---


# Period.Count property (Project)

Gets the number of days in the **Period** object. Read-only **Integer**.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Period](./Project.Period.md) object.


## Example

The following example prompts the user for the name of a resource and then assigns that resource to tasks without any resources.


```vb
Sub AssignResource() 
 
 Dim T As Task ' Task object used in For Each loop 
 Dim R As Resource ' Resource object used in For Each loop 
 Dim Rname As String ' Resource name 
 Dim RID As Long ' Resource ID 
 
 RID = 0 
 RName = InputBox$("Enter the name of a resource: ") 
 
 For Each R in ActiveProject.Resources 
 If R.Name = RName Then 
 RID = R.ID 
 Exit For 
 End If 
 Next R 
 
 If RID <> 0 Then 
 ' Assign the resource to tasks without any resources. 
 For Each T In ActiveProject.Tasks 
 If T.Assignments.Count = 0 Then 
 T.Assignments.Add ResourceID:=RID 
 End If 
 Next T 
 Else 
 MsgBox Prompt:=RName & " is not a resource in this project.", buttons:=vbExclamation 
 End If 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]