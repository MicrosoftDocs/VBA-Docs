---
title: Application.ProjectBeforeTaskNew2 event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeTaskNew2
ms.assetid: 4df0eb83-e60d-943d-aecf-57a2f857ae42
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeTaskNew2 event (Project)

Occurs before one or more tasks are created. Uses the **EventInfo** object parameter.


## Syntax

_expression_. `ProjectBeforeTaskNew2`( `_pj_`, `_Info_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project in which a task or tasks are being created.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is **False** when the event occurs. If the event procedure sets this argument to **True**, the new task or tasks are not created.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

the **ProjectBeforeTaskNew2** event doesn't occur when data is merged or appended into a project, during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]