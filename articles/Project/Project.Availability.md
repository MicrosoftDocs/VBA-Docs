---
title: Availability Object (Project)
ms.prod: project-server
api_name:
- Project.Availability
ms.assetid: 2b832aed-2b58-f020-2a2c-8756ec7ec1a4
ms.date: 06/08/2017
---


# Availability Object (Project)


 

Represents a line from the  **Resource Availability** grid for a resource. The **Availability** object is a member of the **[Availabilities](Project.availabilities.md)** collection.
 
 **Using the Availability Object**
 
Use  **Availabilities(***Index* **)**, where*Index* is the availability index number, to return a single **Availability** object. The following example returns the availability information from the first line of the **Resource Availability** grid for the specified resource.
 



```
MsgBox ActiveProject.Resources("Tom").Name &amp; " is available from " &amp; _ 
    ActiveProject.Resources("Tom").Availabilities(1).AvailableFrom &amp; " to " &amp; _ 
    ActiveProject.Resources("Tom").Availabilities(1).AvailableTo &amp; "." 

```

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
|[Delete](Project.Availability.Delete.md)|

## Properties



|**Name**|
|:-----|
|[Application](Project.Availability.Application.md)|
|[AvailableFrom](Project.Availability.AvailableFrom.md)|
|[AvailableTo](Project.Availability.AvailableTo.md)|
|[AvailableUnit](Project.Availability.AvailableUnit.md)|
|[Index](Project.Availability.Index.md)|
|[Parent](Project.Availability.Parent.md)|

