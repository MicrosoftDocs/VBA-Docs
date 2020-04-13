---
title: Group2 object (Project)
ms.prod: project-server
api_name:
- Project.Group2
ms.assetid: a7a61fa4-e752-006e-a47e-03987b04f01c
ms.date: 06/08/2017
localization_priority: Normal
---


# Group2 object (Project)

Represents a group definition where the group hierarchy can be maintained. A **Group2** object is a member of a **[Groups2](Project.groups2.md)**, **[ResourceGroups2](Project.resourcegroups2(object).md)**, or **[TaskGroups2](Project.taskgroups2(object).md)** collection.
 


## Remarks

The **Group2** object includes the **[MaintainHierarchy](Project.Group2.MaintainHierarchy.md)** property.
 

 
 **Using the Group Object**
 

 
Use  `TaskGroups2(Index)` or `ResourceGroups2(Index)`, where *Index* is the group definition index or group definition name, to return a **Group2** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```vb
ActiveProject.ResourceGroups2("Standard Rate").ShowSummary = True
```


## Methods



|Name|
|:-----|
|[Delete](Project.Group2.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Project.Group2.Application.md)|
|[GroupAssignments](Project.Group2.GroupAssignments.md)|
|[GroupCriteria](Project.Group2.GroupCriteria.md)|
|[Index](Project.Group2.Index.md)|
|[MaintainHierarchy](Project.Group2.MaintainHierarchy.md)|
|[Name](Project.Group2.Name.md)|
|[Parent](Project.Group2.Parent.md)|
|[ShowSummary](Project.Group2.ShowSummary.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]