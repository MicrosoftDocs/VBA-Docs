---
title: Group object (Project)
ms.prod: project-server
api_name:
- Project.Group
ms.assetid: e3756818-f051-1ae4-5402-0398e568ebfc
ms.date: 06/08/2017
localization_priority: Normal
---


# Group object (Project)

Represents a group definition. A **Group** object is a member of the **[ResourceGroups](Project.resourcegroups(object).md)** collection or the **[TaskGroups](Project.taskgroups(object).md)** collection.
 


## Remarks

 **Using the Group Object**
 

 
Use  `TaskGroups(Index)` or `ResourceGroups(Index)`, where *Index* is the group definition index or group definition name, to return a **Group** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```vb
ActiveProject.ResourceGroups("Standard Rate").ShowSummary = True
```


## Methods



|Name|
|:-----|
|[Delete](Project.Group.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Project.Group.Application.md)|
|[GroupAssignments](Project.Group.GroupAssignments.md)|
|[GroupCriteria](Project.Group.GroupCriteria.md)|
|[Index](Project.Group.Index.md)|
|[Name](Project.Group.Name.md)|
|[Parent](Project.Group.Parent.md)|
|[ShowSummary](Project.Group.ShowSummary.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]