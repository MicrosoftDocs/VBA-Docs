---
title: Exception.Shift4 property (Project)
ms.prod: project-server
api_name:
- Project.Exception.Shift4
ms.assetid: 295c1f7c-7614-3efb-4fe6-849c8a2b89d6
ms.date: 06/08/2017
localization_priority: Normal
---


# Exception.Shift4 property (Project)

Gets a  **[Shift](Project.Shift.md)** object representing the fourth work shift in a calendar exception for a day, month, period, weekday, or throughout a year. Read-only **Shift**.


## Syntax

_expression_. `Shift4`

_expression_ A variable that represents an [Exception](./Project.Exception.md) object.


## Example

The following example schedules a half-day of work on Fridays by creating a shift from 8 A.M. to noon.


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


## See also


[Exception Object](Project.Exception.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]