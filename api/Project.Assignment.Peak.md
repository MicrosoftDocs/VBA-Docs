---
title: Assignment.Peak property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Peak
ms.assetid: 52b5d301-6034-b207-c5ae-dfadb56ecd73
ms.date: 06/08/2017
localization_priority: Normal
---


# Assignment.Peak property (Project)

Gets the largest number of resource units for the assignment. Read-only  **Variant**.


## Syntax

_expression_. `Peak`

_expression_ A variable that represents an [Assignment](./Project.Assignment.md) object.


## Example

The following example finds any assignments with more than a certain number of resource units assigned.


```vb
Sub FindOverassigned() 
 Dim T As Task, A As Assignment 
 Dim TooMany As Double, Results As String 
 
 TooMany = InputBox("Enter maximum allowed units per assignment: ") 
 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 For Each A In T.Assignments 
 If A.Peak > TooMany Then 
 Results = Results & T.Name & ": " & A.ResourceName & vbCrLf 
 End If 
 Next A 
 If Results <> "" Then MsgBox "The following resources are " & _ 
 "assigned more than " & TooMany & " units:" & vbCrLf & Results 
 Results = "" 
 End If 
 Next T 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]