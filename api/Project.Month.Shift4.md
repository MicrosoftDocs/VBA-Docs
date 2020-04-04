---
title: Month.Shift4 property (Project)
ms.prod: project-server
api_name:
- Project.Month.Shift4
ms.assetid: 0491c6e9-c02e-59fd-a6d6-e9d2ce88b8c4
ms.date: 06/08/2017
localization_priority: Normal
---


# Month.Shift4 property (Project)

Gets a **[Shift](Project.Shift.md)** object representing the fourth work shift in a month. Read-only **Shift**.


## Syntax

_expression_. `Shift4`

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