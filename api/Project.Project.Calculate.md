---
title: Project.Calculate event (Project)
ms.prod: project-server
api_name:
- Project.Project.Calculate
ms.assetid: cba7feb3-c0e4-96ec-d2fa-eaccfa640c5a
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Calculate event (Project)

Occurs when a project schedule is recalculated.


## Syntax

_expression_. `Calculate`( `_pj_` )

 _expression_ An expression that returns a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that is rescheduled.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]