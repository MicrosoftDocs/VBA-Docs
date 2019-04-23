---
title: WorkWeekDays.Count property (Project)
ms.prod: project-server
api_name:
- Project.WorkWeekDays.Count
ms.assetid: 236d6836-05da-889c-ac76-5876d908e16f
ms.date: 06/08/2017
localization_priority: Normal
---


# WorkWeekDays.Count property (Project)

Gets the number of items in the  **WorkWeekDays** collection. Read-only **Integer**.


## Syntax

_expression_.**Count**

 _expression_ An expression that returns a 'WorkWeekDays' object.


## Example

The following example shows there are seven workweek days in the calendar for the first resource of the active project.


```vb
Debug.Print ActiveProject.Resources(1).Calendar.WorkWeeks(1).WeekDays.Count
```


## See also


[WorkWeekDays Collection Object](Project.workweekdays.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]