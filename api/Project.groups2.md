---
title: Groups2 object (Project)
ms.prod: project-server
ms.assetid: b2b83868-3366-4fb0-fed9-16d4c5eaff87
ms.date: 06/08/2017
localization_priority: Normal
---


# Groups2 object (Project)

Represents a collection of  **[Group2](Project.Group2.md)** objects, which can maintain group hierarchy.
 


## Remarks

Use  `TaskGroups2(Index)` or `ResourceGroups2(Index)`, where *Index* is the group definition index or group definition name, to return a **Group2** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```vb
ActiveProject.ResourceGroups2("Standard Rate").ShowSummary = True 


```


## Methods



|Name|
|:-----|
|[Add](Project.Groups2.Add.md)|
|[Copy](Project.Groups2.Copy.md)|

## Properties



|Name|
|:-----|
|[Application](Project.Groups2.Application.md)|
|[Count](Project.Groups2.Count.md)|
|[Item](Project.Groups2.Item.md)|
|[Parent](Project.Groups2.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]