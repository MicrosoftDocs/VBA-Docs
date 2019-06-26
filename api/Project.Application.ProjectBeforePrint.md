---
title: Application.ProjectBeforePrint event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforePrint
ms.assetid: 7cc8de23-c3e3-81df-ae26-37c4e639dd81
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforePrint event (Project)

Occurs before a project is printed.


## Syntax

_expression_. `ProjectBeforePrint`( `_pj_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**| The project to be printed.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the project will not be printed.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]