---
title: Group.GroupCriteria property (Project)
ms.prod: project-server
api_name:
- Project.Group.GroupCriteria
ms.assetid: c021a7ca-1e80-4318-7612-3d2bf579b683
ms.date: 06/08/2017
localization_priority: Normal
---


# Group.GroupCriteria property (Project)

Gets or sets a **[GroupCriteria](Project.GroupCriterion.md)** collection representing the fields in a group definition. Read/write **GroupCriteria**.


## Syntax

_expression_. `GroupCriteria`

_expression_ A variable that represents a [Group](./Project.Group.md) object.


## Example

The following example adds a criterion to the specified resource group, grouping resources in ascending order as determined by the percentage of their work (in 5% increments) that is complete.


```vb
Sub AddCriterionWithInterval() 
 ActiveProject.ResourceGroups("Response Pending").GroupCriteria.Add "% Work Complete", 
 True, CellColor:=pjRed, GroupOn:=pjGroupOnPctInterval, StartAt:=5, GroupInterval:=5 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]