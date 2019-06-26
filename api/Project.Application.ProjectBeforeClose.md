---
title: Application.ProjectBeforeClose event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeClose
ms.assetid: 90e75c72-03f9-25ab-1339-94d9ff8933a2
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeClose event (Project)

Occurs before a project is closed.


## Syntax

_expression_. `ProjectBeforeClose`( `_pj_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project to be closed|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the project will not be closed.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]