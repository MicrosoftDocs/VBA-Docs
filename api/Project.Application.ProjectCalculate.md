---
title: Application.ProjectCalculate event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectCalculate
ms.assetid: 44dbf3f9-4a7d-2e85-aa63-915ea47af008
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectCalculate event (Project)

Occurs after a project is calculated.


## Syntax

_expression_. `ProjectCalculate`( `_pj_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that was calculated.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]