---
title: WeekDays object (Project)
ms.prod: project-server
ms.assetid: 757437a0-e2ff-0027-f044-87d1cb357f62
ms.date: 06/08/2017
localization_priority: Normal
---


# WeekDays object (Project)

Contains a collection of  **[Weekday](Project.WeekDay.md)** objects.
 


## Example

 **Using the Weekday Object**
 

 
Use  **Weekdays** (*Index* ), where*Index* is the weekday index number, three-letter abbreviation of the day name, or **PjWeekday** constant, to return a single **Weekday** object. The following example sets Friday (the sixth day of a week starting on Sunday) as a half-day by setting the start and finish times for the first shift and clearing the values of the second and third shifts.
 

 



```vb
With ActiveProject.Calendar.WeekDays(6) 

 .Shift1.Start = #8:00:00 AM# 

 .Shift1.Finish = #12:00:00 PM# 

 .Shift2.Clear 

 .Shift3.Clear 

End With
```

A much better way to return the same object is to use the predefined constant for Friday instead of the nonintuitive number 6. Thus, the first line of the preceding example would be as follows: 
 

 



```vb
With ActiveProject.Calendar.WeekDays(pjFriday)
```

 **Using the Weekdays Collection**
 

 
Use the **[Weekdays](Project.Calendar.WeekDays.md)** property to return a **Weekdays** collection.
 

 



```vb
ActiveProject.Calendar.WeekDays
```


## Properties



|Name|
|:-----|
|[Application](Project.WeekDays.Application.md)|
|[Count](Project.WeekDays.Count.md)|
|[Item](Project.WeekDays.Item.md)|
|[Parent](Project.WeekDays.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]