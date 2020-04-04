---
title: Year.Shift3 property (Project)
ms.prod: project-server
api_name:
- Project.Year.Shift3
ms.assetid: eea8a0f6-8889-0d13-f648-e95fc09b2874
ms.date: 06/08/2017
localization_priority: Normal
---


# Year.Shift3 property (Project)

Gets a **[Shift](Project.Shift.md)** object representing the third work shift throughout a year. Read-only **Shift**.


## Syntax

_expression_. `Shift3`

_expression_ A variable that represents a [Year](./Project.Year.md) object.


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