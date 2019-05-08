---
title: Resource.ActualCost property (Project)
ms.prod: project-server
api_name:
- Project.Resource.ActualCost
ms.assetid: 9e5bd065-c88d-aa87-0191-be95b4d3ca04
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.ActualCost property (Project)

Gets the current actual cost for the resource on the project. Read-only  **Variant**.


## Syntax

_expression_. `ActualCost`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Remarks

The current actual cost for the resource is calculated from the resource cost rate tables and the actual work the resource has completed on assignments in the project. For programmatic access to the resource cost rate tables, use the  **[CostRateTables](Project.Resource.CostRateTables.md)** collection.

Actual costs are also available for tasks and assignments. For an example the uses the actual cost for tasks, see the  **[ActualCost](Project.Task.ActualCost.md)** property for the **Task** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]