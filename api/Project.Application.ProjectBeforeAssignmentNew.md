---
title: Application.ProjectBeforeAssignmentNew event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeAssignmentNew
ms.assetid: 5caedd9a-94b1-daa6-762a-a037dae4f917
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeAssignmentNew event (Project)

Occurs before one or more assignments are created.


## Syntax

_expression_. `ProjectBeforeAssignmentNew`( `_pj_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project in which an assignment or assignments are being created.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the new assignment(s) are not created.|

## Return value

**Nothing**


## Remarks

The  **ProjectBeforeAssignmentNew** event also fires when a resource assignment is replaced. Additionally, the event will fire when the only resource assignment on a task is removed, because an "Unassigned Resource" assignment is created after the existing assignment is removed.

Project events do not occur when the project is embedded in another document or application. 

The  **ProjectBeforeAssignmentNew** event doesn't occur when an assignment is created as the result of a drag-and-drop operation in the **Resource Usage** view, during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]