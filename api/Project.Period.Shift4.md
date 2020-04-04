---
title: Period.Shift4 property (Project)
ms.prod: project-server
api_name:
- Project.Period.Shift4
ms.assetid: 64494509-b5dd-2ee3-b933-6a728c50444d
ms.date: 06/08/2017
localization_priority: Normal
---


# Period.Shift4 property (Project)

Gets a **[Shift](Project.Shift.md)** object representing the fourth work shift in a time period. Read-only **Shift**.


## Syntax

_expression_. `Shift4`

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