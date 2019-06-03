---
title: Year object (Project)
keywords: vbapj.chm131361
f1_keywords:
- vbapj.chm131361
ms.prod: project-server
api_name:
- Project.Year
ms.assetid: 060e541f-f709-65dd-c955-5d04c1554373
ms.date: 06/08/2017
localization_priority: Normal
---


# Year object (Project)

Represents a year in a project calendar. The  **Year** object is a member of the **[Years](Project.years.md)** collection.
 


## Remarks

The  **Years** collection in Project begins in 1984 and ends in 2149. In previous versions of Project, scheduling can run from 1984 to 2049.
 

 

## Examples

 **Using the Year Object**
 

 
Use  **Years** ( _Index_), where  _Index_ is the year index number, to return a single **Year** object. The following example counts the number of working days in the month of September 2012 for each selected resource.
 

 



```vb
Dim r As Resource
Dim d As Integer
Dim workingDays As Integer
Dim theMonth As PjMonth

theMonth = pjSeptember

For Each r In ActiveSelection.Resources()
    workingDays = 0
    With r.Calendar.Years(2012).Months(theMonth)
        For d = 1 To .Days.Count
            If .Days(d).Working = True Then
                workingDays = workingDays + 1
            End If
        Next d
    End With
    MsgBox "There are " & workingDays & " working days in " _
        & r.Name & "'s calendar for month " & theMonth
Next r
```

 **Using the Years Collection**
 

 
Use the  **[Years](Project.Calendar.Years.md)** property to return a **Years** collection. The following example lists all the years in the calendar of the active project.
 

 



```vb
Sub CountYears()
    Dim c As Long
    Dim temp As String
        
    For c = 1 To ActiveProject.Calendar.Years.Count
        temp = temp & ListSeparator & " " & _
            ActiveProject.Calendar.Years(c + 1983).Name
    Next c
            
    MsgBox Right$(temp, Len(temp) - Len(ListSeparator & " "))
End Sub
```

Figure 1 shows the results of the  **CountYears** macro.
 

 

**Figure 1. Getting the list of years available**

 
![Years available for project planning](../images/pj15_VBA_Years.gif)
 

 

## Methods



|Name|
|:-----|
|[Default](Project.Year.Default.md)|

## Properties



|Name|
|:-----|
|[Application](Project.Year.Application.md)|
|[Calendar](Project.Year.Calendar.md)|
|[Count](Project.Year.Count.md)|
|[Index](Project.Year.Index.md)|
|[Months](Project.Year.Months.md)|
|[Name](Project.Year.Name.md)|
|[Parent](Project.Year.Parent.md)|
|[Shift1](Project.Year.Shift1.md)|
|[Shift2](Project.Year.Shift2.md)|
|[Shift3](Project.Year.Shift3.md)|
|[Shift4](Project.Year.Shift4.md)|
|[Shift5](Project.Year.Shift5.md)|
|[Working](Project.Year.Working.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]