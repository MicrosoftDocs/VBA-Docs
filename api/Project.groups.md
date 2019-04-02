---
title: Groups object (Project)
ms.prod: project-server
ms.assetid: 2e4c4846-6193-fc12-ad02-0dd69f88b31e
ms.date: 06/08/2017
localization_priority: Normal
---


# Groups object (Project)

Represents a collection of  **[Group](Project.Group.md)** objects.
 


## Remarks

For groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, use the  **[Groups2](Project.groups2.md)** collection object.
 

 
Use  `TaskGroups(Index)` or `ResourceGroups(Index)`, where *Index* is the group definition index or group definition name, to return a **Group** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```vb
ActiveProject.ResourceGroups("Standard Rate").ShowSummary = True 


```


## Methods



|Name|
|:-----|
|[Add](Project.Groups.Add.md)|
|[Copy](Project.Groups.Copy.md)|

## Properties



|Name|
|:-----|
|[Application](Project.Groups.Application.md)|
|[Count](Project.Groups.Count.md)|
|[Item](Project.Groups.Item.md)|
|[Parent](Project.Groups.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]