---
title: Task.LinkSuccessors method (Project)
ms.prod: project-server
api_name:
- Project.Task.LinkSuccessors
ms.assetid: 397fff8c-3ff3-4725-2938-fdaecddf624b
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.LinkSuccessors method (Project)

Adds one or more successors to the task.


## Syntax

_expression_. `LinkSuccessors`( `_Tasks_`, `_Link_`, `_Lag_` )

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Tasks_|Required|**Object**|The **Task** or **Tasks** object specified becomes a successor of the task specified with **expression**.|
| _Link_|Optional|**Long**| A constant that specifies the relationship between tasks that become linked. Can be one of the [PjTaskLinkType](Project.PjTaskLinkType.md) constants. The default value is **pjFinishToStart**.|
| _Lag_|Optional|**Variant**|A string that specifies the duration of lag time between linked tasks. To specify lead time between tasks, use an expression for  **Lag** that evaluates to a negative value.|

## Return value

Nothing


## Example

The following example create two tasks and links the second task as successor to the first.


```vb
Sub Link_Successors() 
    Dim SuccessorTask As Task 
    Dim PredecessorTask As Task 
 
    'Activate Task Sheet view 
    ViewApply Name:="Task Sheet" 
 
    ' Create a coupe of tasks 
    RowInsert 
    SetTaskField Field:="Name", Value:="TestTask-2" 
    SetTaskField Field:="Duration", Value:="1" 
 
    RowInsert 
    SetTaskField Field:="Name", Value:="TestTask-1" 
    SetTaskField Field:="Duration", Value:="2" 
 
    'link them 
    Set PredecessorTask = ActiveProject.Tasks("TestTask-1") 
    Set SuccessorTask = ActiveProject.Tasks("TestTask-2") 
 
    PredecessorTask.LinkSuccessors Tasks:=SuccessorTask, Link:=pjFinishToStart 
 
    'delete the tasks 
    PredecessorTask.Delete 
    SuccessorTask.Delete 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]