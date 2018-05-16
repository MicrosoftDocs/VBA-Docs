---
title: Availabilities Object (Project)
ms.prod: project-server
ms.assetid: 51224d62-777b-1ae3-a646-ca977464d37d
ms.date: 06/08/2017
---


# Availabilities Object (Project)

 Contains a collection of **[Availability](Project.Availability.md)** objects.
 


## Example

 **Using the Availabilities Collection**
 

 
Use  **Availabilities(** Index **)**, where Index is the availability index number, to return a single **Availability** object. The following example returns the availability information from the first line of the **Resource Availability** grid for the specified resource.
 

 



```
MsgBox ActiveProject.Resources("Tom").Name &amp; " is available from " &amp; _  
    ActiveProject.Resources("Tom").Availabilities(1).AvailableFrom &amp; " to " &amp; _  
    ActiveProject.Resources("Tom").Availabilities(1).AvailableTo &amp; "."  

```

 **Using the Availabilities Collection**
 

 
Use the  **[Availabilities](Project.Resource.Availabilities.md)** property to return an **Availabilities** collection. The following example displays the range of dates during which the specified resource is available for work.
 

 



```
Dim Avail As Availability  

For Each Avail In ActiveProject.Resources("Tom").Availabilities  
    MsgBox "From " &amp; Avail.AvailableFrom &amp; " to " &amp; Avail.AvailableTo  
Next Avail
```

Use the  **[Add](Project.Availabilities.Add.md)** method to add an **Availability** object to the **Availabilities** collection. The following example adds a line to the **Resource Availability** grid showing that the specified resource is available only half-time during the month of April.
 

 



```
ActiveProject.Resources("Tom").Availabilities.Add "4/1/2012", "4/30/2012", 50
```


## Methods



|**Name**|
|:-----|
|[Add](Project.Availabilities.Add.md)|

## Properties



|**Name**|
|:-----|
|[Application](Project.Availabilities.Application.md)|
|[Count](Project.Availabilities.Count.md)|
|[Item](Project.Availabilities.Item.md)|
|[Parent](availabilities-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
