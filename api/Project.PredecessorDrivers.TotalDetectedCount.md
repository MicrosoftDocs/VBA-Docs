---
title: PredecessorDrivers.TotalDetectedCount property (Project)
ms.prod: project-server
api_name:
- Project.PredecessorDrivers.TotalDetectedCount
ms.assetid: 479cc962-5156-6f30-b304-5f4a6bc3abea
ms.date: 06/08/2017
localization_priority: Normal
---


# PredecessorDrivers.TotalDetectedCount property (Project)

Gets the total number of predecessor tasks that affect the start date of a task. Read-only  **Long**.


## Syntax

_expression_. `TotalDetectedCount`

_expression_ A variable that represents a 'PredecessorDrivers' object.


## Remarks

Predecessor tasks are tasks that are linked to the current task and occur before it. Predecessor tasks can have constraints and lag or lead time and can themselves have other predecessors that affect the total count of predecessor drivers.


## See also


[PredecessorDrivers Collection Object](Project.predecessordrivers.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]