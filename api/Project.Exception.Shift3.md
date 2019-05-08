---
title: Exception.Shift3 property (Project)
ms.prod: project-server
api_name:
- Project.Exception.Shift3
ms.assetid: e7a320de-d140-81f9-e0ff-6bf6c4452c31
ms.date: 06/08/2017
localization_priority: Normal
---


# Exception.Shift3 property (Project)

Gets a  **[Shift](Project.Shift.md)** object representing the third work shift in a calendar exception for a day, month, period, weekday, or throughout a year. Read-only **Shift**.


## Syntax

_expression_. `Shift3`

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