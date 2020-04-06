---
title: Period.Shift5 property (Project)
ms.prod: project-server
api_name:
- Project.Period.Shift5
ms.assetid: 4e3e363f-cd36-77b0-4572-4e43e8b58bd5
ms.date: 06/08/2017
localization_priority: Normal
---


# Period.Shift5 property (Project)

Gets a **[Shift](Project.Shift.md)** object representing the fifth work shift in a time period. Read-only **Shift**.


## Syntax

_expression_. `Shift5`

_expression_ A variable that represents a [Period](./Project.Period.md) object.


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