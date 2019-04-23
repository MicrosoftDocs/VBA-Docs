---
title: WeekDay object (Project)
ms.prod: project-server
api_name:
- Project.WeekDay
ms.assetid: fc460e89-784b-6764-c22d-e1dcd8a9f297
ms.date: 06/08/2017
localization_priority: Normal
---


# WeekDay object (Project)


 

Represents a weekday in a calendar. The  **Weekday** object is a member of the **[Weekdays](Project.weekdays.md)** collection.
 
 **Using the Weekday Object**
 
Use  **Weekdays** (*Index* ), where*Index* is the weekday index number, three-letter abbreviation of the day name, or **PjWeekday** constant, to return a single **Weekday** object. The following example sets Friday (the sixth day of a week starting on Sunday) as a half-day by setting the start and finish times for the first shift and clearing the values of the second and third shifts.
 
A much better way to return the same object is to use the predefined constant for Friday instead of the nonintuitive number 6. Thus, the first line of the preceding example would be as follows:
 
 **Using the Weekdays Collection**
 
Use the  **[Weekdays](Project.Calendar.WeekDays.md)** property to return a **Weekdays** collection.
 

## Methods



|Name|
|:-----|
|[Default](Project.WeekDay.Default.md)|

## Properties



|Name|
|:-----|
|[Application](Project.WeekDay.Application.md)|
|[Calendar](Project.WeekDay.Calendar.md)|
|[Count](Project.WeekDay.Count.md)|
|[Index](Project.WeekDay.Index.md)|
|[Name](Project.WeekDay.Name.md)|
|[Parent](Project.WeekDay.Parent.md)|
|[Shift1](Project.WeekDay.Shift1.md)|
|[Shift2](Project.WeekDay.Shift2.md)|
|[Shift3](Project.WeekDay.Shift3.md)|
|[Shift4](Project.WeekDay.Shift4.md)|
|[Shift5](Project.WeekDay.Shift5.md)|
|[Working](Project.WeekDay.Working.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]