---
title: WeekDay.Shift5 property (Project)
ms.prod: project-server
api_name:
- Project.WeekDay.Shift5
ms.assetid: 1bfee704-e2cd-8fdd-23a4-a182c853dca3
ms.date: 06/08/2017
localization_priority: Normal
---


# WeekDay.Shift5 property (Project)

Gets a **[Shift](Project.Shift.md)** object representing the fifth work shift in a weekday. Read-only **Shift**.


## Syntax

_expression_. `Shift5`

_expression_ A variable that represents a [WeekDay](./Project.WeekDay.md) object.


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