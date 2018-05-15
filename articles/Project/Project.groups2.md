---
title: Groups2 Object (Project)
ms.prod: project-server
ms.assetid: b2b83868-3366-4fb0-fed9-16d4c5eaff87
ms.date: 06/08/2017
---


# Groups2 Object (Project)

Represents a collection of  **[Group2](Project.Group2.md)** objects, which can maintain group hierarchy.
 


## Remarks

Use  `TaskGroups2(Index)` or `ResourceGroups2(Index)`, where *Index* is the group definition index or group definition name, to return a **Group2** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```
ActiveProject.ResourceGroups2("Standard Rate").ShowSummary = True 


```


## Methods



|**Name**|
|:-----|
|[Add](Project.Groups2.Add.md)|
|[Copy](Project.Groups2.Copy.md)|

## Properties



|**Name**|
|:-----|
|[Application](Project.Groups2.Application.md)|
|[Count](Project.Groups2.Count.md)|
|[Item](Project.Groups2.Item.md)|
|[Parent](Project.Groups2.Parent.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
