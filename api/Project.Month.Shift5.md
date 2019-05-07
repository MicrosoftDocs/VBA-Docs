---
title: Month.Shift5 property (Project)
ms.prod: project-server
api_name:
- Project.Month.Shift5
ms.assetid: 662b3249-a68a-7fe0-4d45-71e5fa4c9ec1
ms.date: 06/08/2017
localization_priority: Normal
---


# Month.Shift5 property (Project)

Gets a  **[Shift](Project.Shift.md)** object representing the fifth work shift in a month. Read-only **Shift**.


## Syntax

_expression_. `Shift5`

_expression_ A variable that represents a [Month](./Project.Month.md) object.


## Example

The following example schedules a half-day of work on Fridays by creating an 8 A.M. to noon shift.


```vb
Sub HalfDayFridays() 
 
 With ActiveProject.Calendar.WeekDays(pjFriday) 
 .Shift1.Start = #8:00:00 AM# 
 .Shift1.Finish = #12:00:00 PM# 
 .Shift2.Clear 
 .Shift3.Clear 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]