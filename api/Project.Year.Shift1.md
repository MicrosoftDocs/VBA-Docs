---
title: Year.Shift1 property (Project)
ms.prod: project-server
api_name:
- Project.Year.Shift1
ms.assetid: 4c352439-21c1-e369-7a33-d8e92ba23f2d
ms.date: 06/08/2017
localization_priority: Normal
---


# Year.Shift1 property (Project)

Gets a **[Shift](Project.Shift.md)** object representing the first work shift throughout a year. Read-only **Shift**.


## Syntax

_expression_. `Shift1`

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