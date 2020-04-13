---
title: WorkWeekDay.Count property (Project)
ms.prod: project-server
api_name:
- Project.WorkWeekDay.Count
ms.assetid: 242bb040-d7ec-187f-4946-c5d38c8c29a0
ms.date: 06/08/2017
localization_priority: Normal
---


# WorkWeekDay.Count property (Project)

Gets the value 1 for the number of days in the **WorkWeekDay** object. Read-only **Integer**.


## Syntax

_expression_.**Count**

 _expression_ An expression that returns a [WorkWeekDay](./Project.WorkWeekDay.md) object.


## Example

The following example shows there is one day in the fourth day of the work week.


```vb
Debug.Print ActiveProject.Resources(1).Calendar.WorkWeeks(1).WeekDays(4).Count
```


## See also


[WorkWeekDay Object](Project.WorkWeekDay.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]