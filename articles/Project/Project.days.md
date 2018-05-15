---
title: Days Object (Project)
ms.prod: project-server
ms.assetid: ac9cc007-a318-c9a8-2e6c-c4834a52d5c2
ms.date: 06/08/2017
---


# Days Object (Project)

Contains a collection of  **[Day](Project.Day.md)** objects.
 


## Example

 **Using the Days Collection Object**
 

 
Use  **Days(***Index* **)**, where*Index* is the day index number or **[PjWeekday](Project.PjWeekday.md)** constant, to return a single **Day** object. The following example counts the number of working days in the month of September 2002 for each selected resource.
 

 



```
Dim R As Resource, D As Integer, WorkingDays As Integer 

 

For Each R In ActiveSelection.Resources() 

 WorkingDays = 0 

 With R.Calendar.Years(2002).Months(pjSeptember) 

 For D = 1 To .Days.Count 

 If .Days(D).Working = True Then 

 WorkingDays = WorkingDays + 1 

 End If 

 Next D 

 End With 

 MsgBox "There are " &amp; WorkingDays &amp; " working days in " _ 

 &amp; R.Name &amp; "'s calendar." 

Next R
```

 **Getting the Days Collection Object.**
 

 
Use the  **[Days](Project.Month.Days.md)** property to return a **Days** collection. The following example counts the number of days in the month of September 2002.
 

 



```
MsgBox ActiveProject.Calendar.Years(2006).Months(pjNovember).Days.Count 


```


## Properties



|**Name**|
|:-----|
|[Application](Project.Days.Application.md)|
|[Count](Project.Days.Count.md)|
|[Item](Project.Days.Item.md)|
|[Parent](Project.Days.Parent.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
