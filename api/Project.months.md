---
title: Months object (Project)
ms.prod: project-server
ms.assetid: 5db0ed37-cc23-7bc8-ebe5-fdaf6275b5db
ms.date: 06/08/2017
localization_priority: Normal
---


# Months object (Project)

Contains a collection of  **[Month](Project.Month.md)** objects.
 


## Remarks

Use  **Months** (*Index* ), where*Index* is the month index number, month name, or **PjMonth** constant, to return a single **Month** object.
 

 

## Example

 **Using the Months Collection Object**
 

 
The following example counts the number of working days in each month of 2012 for each selected resource. 
 

 



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
 

 
Use the **[Months](Project.Year.Months.md)** property to return a **Months** collection. The following example counts the number of months in 2012.
 

 



```vb
ActiveProject.Calendar.Years(2012).Months.Count
```


## Properties



|Name|
|:-----|
|[Application](Project.Months.Application.md)|
|[Count](Project.Months.Count.md)|
|[Item](Project.Months.Item.md)|
|[Parent](Project.Months.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]