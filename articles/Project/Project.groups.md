---
title: Groups Object (Project)
ms.prod: project-server
ms.assetid: 2e4c4846-6193-fc12-ad02-0dd69f88b31e
ms.date: 06/08/2017
---


# Groups Object (Project)

Represents a collection of  **[Group](Project.Group.md)** objects.
 


## Remarks

For groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, use the  **[Groups2](Project.groups2.md)** collection object.
 

 
Use  `TaskGroups(Index)` or ` ResourceGroups(Index)`, where *Index* is the group definition index or group definition name, to return a **Group** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```
ActiveProject.ResourceGroups("Standard Rate").ShowSummary = True 


```


## Methods



|**Name**|
|:-----|
|[Add](Project.Groups.Add.md)|
|[Copy](Project.Groups.Copy.md)|

## Properties



|**Name**|
|:-----|
|[Application](Project.Groups.Application.md)|
|[Count](Project.Groups.Count.md)|
|[Item](Project.Groups.Item.md)|
|[Parent](Project.Groups.Parent.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
