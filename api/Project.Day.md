---
title: Day object (Project)
ms.prod: project-server
api_name:
- Project.Day
ms.assetid: 411fe04f-b68d-08c2-8b6c-f2c1e9927a34
ms.date: 06/08/2017
localization_priority: Normal
---


# Day object (Project)

Represents a day in a month. The **Day** object is a member of the **[Days](Project.days.md)** collection.
 


## Example

 **Using the Day Object**
 

 
Use  **Days** (*Index* ), where*Index* is the day index number or **[PjWeekday](Project.PjWeekday.md)** constant, to return a single **Day** object. The following example counts the number of working days in the month of September 2008 for each selected resource.
 

 



```vb
Dim R As Resource, D As Integer, WorkingDays As Integer 
 
For Each R In ActiveSelection.Resources() 
    WorkingDays = 0 
    With R.Calendar.Years(2008).Months(pjSeptember) 
        For D = 1 To .Days.Count 
            If .Days(D).Working = True Then 
                WorkingDays = WorkingDays + 1 
            End If 
        Next D 
    End With 
    MsgBox "There are " & WorkingDays & " working days in " _ 
        & R.Name & "'s calendar." 
Next R
```

 **Using the Days Collection**
 

 
Use the **[Days](Project.Month.Days.md)** property to return a **Days** collection. The following example counts the number of working days in the month of September 2008.
 

 



```vb
ActiveProject.Calendar.Years(2008).Months(pjSeptember).Days.Count
```


## Methods



|Name|
|:-----|
|[Default](Project.Day.Default.md)|

## Properties



|Name|
|:-----|
|[Application](Project.Day.Application.md)|
|[Calendar](Project.Day.Calendar.md)|
|[Count](Project.Day.Count.md)|
|[Index](Project.Day.Index.md)|
|[Name](Project.Day.Name.md)|
|[Parent](Project.Day.Parent.md)|
|[Shift1](Project.Day.Shift1.md)|
|[Shift2](Project.Day.Shift2.md)|
|[Shift3](Project.Day.Shift3.md)|
|[Shift4](Project.Day.Shift4.md)|
|[Shift5](Project.Day.Shift5.md)|
|[Working](Project.Day.Working.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]