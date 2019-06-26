---
title: Application.ProjectAssignmentNew event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectAssignmentNew
ms.assetid: dcb4acc6-a113-1e93-5f08-e9e68b902b96
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectAssignmentNew event (Project)

Occurs when a new assignment is created.


## Syntax

_expression_. `ProjectAssignmentNew`( `_pj_`, `_ID_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project where the assignment was just created.|
| _ID_|Required|**Long**|The ID of the assignment that was just created.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]