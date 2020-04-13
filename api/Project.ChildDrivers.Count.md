---
title: ChildDrivers.Count property (Project)
ms.prod: project-server
api_name:
- Project.ChildDrivers.Count
ms.assetid: 6d5f72e2-b563-84d0-ae14-54fddb32c20e
ms.date: 06/08/2017
localization_priority: Normal
---


# ChildDrivers.Count property (Project)

Gets the number of items in the **[ChildDrivers](Project.childdrivers.md)** collection. Read-only **Long**.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a 'ChildDrivers' object.


## Remarks

If  **TotalDetectedCount** is greater than 5 then count is 0.

Use of the **Count** property in most collection objects is similar. For an example, see the **[Assignments.Count](Project.Assignments.Count.md)** property.


## See also


[ChildDrivers Collection Object](Project.childdrivers.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]