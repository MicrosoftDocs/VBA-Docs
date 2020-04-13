---
title: Groups2.Copy method (Project)
ms.prod: project-server
api_name:
- Project.Groups2.Copy
ms.assetid: a0b45d11-394a-4915-5eb8-62ffaab04757
ms.date: 06/08/2017
localization_priority: Normal
---


# Groups2.Copy method (Project)

Makes a copy of a group definition from the **Groups2** collection and returns a reference to the **[Group2](Project.Group2.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _NewName_)

 _expression_ An expression that returns a 'Groups2' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

## Return value

 **Group2**


## See also


[Groups2 Collection Object](Project.groups2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]