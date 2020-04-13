---
title: StartDriver.OverAllocatedAssignments property (Project)
ms.prod: project-server
api_name:
- Project.StartDriver.OverAllocatedAssignments
ms.assetid: bef55fa0-e721-27f6-aa3b-6314aeaef0fa
ms.date: 06/08/2017
localization_priority: Normal
---


# StartDriver.OverAllocatedAssignments property (Project)

Gets overallocated assignments for a task start driver. Read-only  **OverAllocatedAssignments**.


## Syntax

_expression_. `OverAllocatedAssignments`( `_fOverPeak_` )

 _expression_ An expression that returns a [StartDriver](./Project.StartDriver.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _overallocationType_|Required|**PjOverallocationType**|Can be one of the **[PjOverallocationType](Project.PjOverallocationType.md)** constants, which determines the type of overallocation.|

## Remarks

Overallocated assignments are not possible on milestones, placeholder tasks, or tasks with no assignments.


## Example

The following command returns the number of overallocated assignments where resources are working on other tasks.


```vb
Debug.Print ActiveProject.Tasks(2).StartDriver.OverAllocatedAssignments(pjOverallocationTypeWorkingOnOtherTasks).Count
```


## See also


[StartDriver Object](Project.StartDriver.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]