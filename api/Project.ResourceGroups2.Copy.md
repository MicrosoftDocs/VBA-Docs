---
title: ResourceGroups2.Copy method (Project)
ms.prod: project-server
api_name:
- Project.ResourceGroups2.Copy
ms.assetid: 3de6fbeb-9067-5ab1-590e-82d2d3c9a136
ms.date: 06/08/2017
localization_priority: Normal
---


# ResourceGroups2.Copy method (Project)

Makes a copy of a group definition for the **ResourceGroups2** collection and returns a reference to the **[Group2](Project.Group2.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _NewName_)

 _expression_ An expression that returns a 'ResourceGroups2' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

## Return value

 **Group2**


## See also


[ResourceGroups2 Collection Object](Project.resourcegroups2(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]