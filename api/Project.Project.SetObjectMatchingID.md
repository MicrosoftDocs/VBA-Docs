---
title: Project.SetObjectMatchingID method (Project)
ms.prod: project-server
api_name:
- Project.Project.SetObjectMatchingID
ms.assetid: d0d79e0a-bfec-9882-bfe9-72f7c51f0baf
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.SetObjectMatchingID method (Project)

Sets the matching identification value of an object in the **Organizer** dialog box, for example to change the view specified by "Gantt Chart".


## Syntax

_expression_. `SetObjectMatchingID`( `_ObjectType_`, `_ObjectName_`, `_MatchingID_` )

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**Long**|The type of object, specified by a **[pjOrganizer](Project.PjOrganizer.md)** constant.|
| _ObjectName_|Required|**String**|Display name of the object.|
| _MatchingID_|Required|**String**|String specifying the matching ID to set.|

## Example

The following example sets the matching ID of a **pjView** object type with the display name "Gantt Chart" to "Gantt Chart 1".


```vb
ActiveProject.SetObjectMatchingID ObjectType:=pjView, ObjectName:="Gantt Chart", MatchingID:="Gantt Chart 1"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]