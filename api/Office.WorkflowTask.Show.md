---
title: WorkflowTask.Show method (Office)
keywords: vbaof11.chm280010
f1_keywords:
- vbaof11.chm280010
ms.prod: office
api_name:
- Office.WorkflowTask.Show
ms.assetid: a7256356-c935-e9ce-e510-6798ebd5563f
ms.date: 01/29/2019
localization_priority: Normal
---


# WorkflowTask.Show method (Office)

Displays a workflow task edit user interface for the specified **WorkflowTask** object.


## Syntax

_expression_.**Show**

_expression_ An expression that returns a **[WorkflowTask](Office.WorkflowTask.md)** object.


## Return value

Integer


## Example

The following example displays the name of each workflow task in the current document and then displays the workflow task edit user interface for a specific task.


```vb
Sub DisplayWorkTask() 
Dim objWorkflowTasks As WorkflowTasks 
Dim objWorkflowTask As WorkflowTask 
Dim cnt As Integer 
 
Set objWorkflowTasks = Document.GetWorkflowTasks() 
 
For cnt = 1 To objWorkflowTasks.Count 
 Debug.Print objWorkflowTask(cnt).Name 
Next 
 
Set objWorkflowTask = objWorkflowTasks(1) 
objWorkflowTask.Show 
 
End Sub 

```


## See also

- [WorkflowTask object members](overview/Library-Reference/workflowtask-members-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]