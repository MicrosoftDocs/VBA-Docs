---
title: WeekDay.Count property (Project)
ms.prod: project-server
api_name:
- Project.WeekDay.Count
ms.assetid: 91828803-9d2f-a7ea-f917-f1e26147f177
ms.date: 06/08/2017
localization_priority: Normal
---


# WeekDay.Count property (Project)

Gets the value 1 for the number of days in the **WeekDay** object. Read-only **Integer**.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [WeekDay](./Project.WeekDay.md) object.


## Example

The following example shows there is one day in the third day of the work week.


```vb
Debug.Print ActiveProject.Resources(1).Calendar.WorkWeeks(1).WeekDays(3).Count
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]