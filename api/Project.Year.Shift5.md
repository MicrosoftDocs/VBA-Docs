---
title: Year.Shift5 property (Project)
ms.prod: project-server
api_name:
- Project.Year.Shift5
ms.assetid: 5b076a75-7576-5f52-ed90-3615cb041e07
ms.date: 06/08/2017
localization_priority: Normal
---


# Year.Shift5 property (Project)

Gets a **[Shift](Project.Shift.md)** object representing the fifth work shift throughout a year. Read-only **Shift**.


## Syntax

_expression_. `Shift5`

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