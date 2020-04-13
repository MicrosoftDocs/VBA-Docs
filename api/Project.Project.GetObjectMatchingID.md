---
title: Project.GetObjectMatchingID method (Project)
keywords: vbapj.chm132294
f1_keywords:
- vbapj.chm132294
ms.prod: project-server
api_name:
- Project.Project.GetObjectMatchingID
ms.assetid: 6e20f9a9-2090-6ea5-e476-70652e866cdf
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.GetObjectMatchingID method (Project)

Returns the matching identification name of an object.


## Syntax

_expression_. `GetObjectMatchingID`( `_ObjectType_`, `_ObjectName_` )

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**Long**|The type of object. Can be one of the **[PjOrganizer](Project.PjOrganizer.md)** constants.|
| _ObjectName_|Required|**String**|Display name of the object.|

## Return value

 **String**


## Remarks

You can use the **GetObjectMatchingID** method to get the display name in a multilanguage environment that uses Object Matching Identifier (OMID) fields. For more information, see the **[UseOMIDs](Project.Application.UseOMIDs.md)** property.


## Example

The following example gets the Matching ID of a View object with the display name "My Gantt Chart".


```vb
MsgBox(ActiveProject.GetObjectMatchingID(pjView, "Gantt Chart"))
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]