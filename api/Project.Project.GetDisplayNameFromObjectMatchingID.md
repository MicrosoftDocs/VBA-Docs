---
title: Project.GetDisplayNameFromObjectMatchingID method (Project)
ms.prod: project-server
api_name:
- Project.Project.GetDisplayNameFromObjectMatchingID
ms.assetid: 5e535f7b-fbd9-2989-57ed-583f491a448b
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.GetDisplayNameFromObjectMatchingID method (Project)

Returns the display name of an object.


## Syntax

_expression_. `GetDisplayNameFromObjectMatchingID`( `_ObjectType_`, `_MatchingID_` )

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**Long**|The type of object. Can be one of the **[PjOrganizer](Project.PjOrganizer.md)** constants.|
| _MatchingID_|Required|**String**|String specifying the matching name of the object.|

## Return value

 **String**


## Remarks

You can use the **GetDisplayNameFromObjectMatchingID** method to get the display name in a multilanguage environment that uses Object Matching Identifier (OMID) fields. For more information, see the **[UseOMIDs](Project.Application.UseOMIDs.md)** property.


## Example

The following example gets the display name of View object that has the Matching ID "My Gantt Chart".


```vb
MsgBox(ActiveProject.GetDisplayNameFromObjectMatchingID(pjView, "My Gantt Chart"))
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]