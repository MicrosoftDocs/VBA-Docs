---
title: Groups.Copy method (Project)
ms.prod: project-server
api_name:
- Project.Groups.Copy
ms.assetid: fa53fb17-be05-ab03-c08b-a2c9034b7da6
ms.date: 06/08/2017
localization_priority: Normal
---


# Groups.Copy method (Project)

Makes a copy of a group definition for the  **Groups** collection and returns a reference to the **[Group](Project.Group.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _NewName_)

_expression_ A variable that represents a 'Groups' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

## Return value

 **Group**


## See also


[Groups Collection Object](Project.groups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]