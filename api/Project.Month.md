---
title: Month object (Project)
ms.prod: project-server
api_name:
- Project.Month
ms.assetid: 5ee32f12-72aa-fa16-ead2-97949005cd7c
ms.date: 06/08/2017
localization_priority: Normal
---


# Month object (Project)

Represents a month in a year. The  **Month** object is a member of the **[Months](Project.months.md)** collection.
 


## Example

 **Using the Month Object**
 

 
Use  **Months** (*Index* ), where*Index* is the month index number, month name, or **PjMonth** constant, to return a single **Month** object. The following example counts the number of working days in each month of 2012 for each selected resource.
 

 



```vb
Dim R As Resource 
Dim D As Integer, M As Integer, WorkingDays As Integer 
 
For Each R In ActiveSelection.Resources() 
    WorkingDays = 0 

    With R.Calendar.Years(2012) 
        For M = 1 To .Months.Count 
            WorkingDays = 0 
            For D = 1 To .Months(M).Days.Count 
                If .Months(M).Days(D).Working = True Then 
                    WorkingDays = WorkingDays + 1 
                End If 
            Next D 

            MsgBox "There are " & WorkingDays & " working days in " & _
                .Months(M).Name & " for " & R.Name & "." 
        Next M 
    End With 
Next R
```

 **Using the Months Collection**
 

 
Use the  **[Months](Project.Year.Months.md)** property to return a **Months** collection. The following example counts the number of months in 2012.
 

 



```vb
ActiveProject.Calendar.Years(2012).Months.Count
```


## Methods



|Name|
|:-----|
|[Default](Project.Month.Default.md)|

## Properties



|Name|
|:-----|
|[Application](Project.Month.Application.md)|
|[Calendar](Project.Month.Calendar.md)|
|[Count](Project.Month.Count.md)|
|[Days](Project.Month.Days.md)|
|[Index](Project.Month.Index.md)|
|[Name](Project.Month.Name.md)|
|[Parent](Project.Month.Parent.md)|
|[Shift1](Project.Month.Shift1.md)|
|[Shift2](Project.Month.Shift2.md)|
|[Shift3](Project.Month.Shift3.md)|
|[Shift4](Project.Month.Shift4.md)|
|[Shift5](Project.Month.Shift5.md)|
|[Working](Project.Month.Working.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]