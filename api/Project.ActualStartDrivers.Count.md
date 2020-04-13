---
title: ActualStartDrivers.Count property (Project)
ms.prod: project-server
api_name:
- Project.ActualStartDrivers.Count
ms.assetid: 57301614-c781-1504-eb99-95ca6a4cdcc6
ms.date: 06/08/2017
localization_priority: Normal
---


# ActualStartDrivers.Count property (Project)

Gets the number of  **Assignment** objects in the **ActualStartDrivers** collection. Read-only **Long**.


## Syntax

_expression_.**Count**

_expression_ A variable that represents an 'ActualStartDrivers' object.


## Remarks

This property returns a read-only  **Long** value in the range 0 through 5; if **TotalDetectedCount** is greater than 5, **Count** returns 0.

Use of the **Count** property in most collection objects is similar. For an example, see the **[Assignments.Count](Project.Assignments.Count.md)** property.


## See also


[ActualStartDrivers Collection Object](Project.actualstartdrivers.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]