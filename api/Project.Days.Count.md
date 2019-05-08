---
title: Days.Count property (Project)
ms.prod: project-server
api_name:
- Project.Days.Count
ms.assetid: 437cc8a8-aa3d-06f1-6327-2830e87e5710
ms.date: 06/08/2017
localization_priority: Normal
---


# Days.Count property (Project)

Gets the number of items in the  **Days** collection. Read-only **Integer**.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a 'Days' object.


## Remarks

Use of the  **Count** property in most collection objects is similar. For an example, see the **[Assignments.Count](Project.Assignments.Count.md)** property.


## Example

The following example shows there are seven days in the  **WeekDays** collection for a resource calendar.


```vb
Debug.Print ActiveProject.Resources(1).Calendar.WeekDays.Count
```


## See also


[Days Collection Object](Project.days.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]