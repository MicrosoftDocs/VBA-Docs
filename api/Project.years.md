---
title: Years object (Project)
ms.prod: project-server
ms.assetid: 3aa139cf-2fc2-7039-5659-8e2d833b5a4f
ms.date: 06/08/2017
localization_priority: Normal
---


# Years object (Project)

Contains a collection of  **[Year](Project.Year.md)** objects.
 


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
 

 

## Properties



|Name|
|:-----|
|[Application](Project.Years.Application.md)|
|[Count](Project.Years.Count.md)|
|[Item](Project.Years.Item.md)|
|[Parent](Project.Years.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]