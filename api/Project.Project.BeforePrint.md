---
title: Project.BeforePrint event (Project)
ms.prod: project-server
api_name:
- Project.Project.BeforePrint
ms.assetid: df66b52b-4c7b-e3e1-d8ff-66416edcb378
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.BeforePrint event (Project)

Occurs before a project is printed.


## Syntax

_expression_. `BeforePrint`( `_pj_` )

 _expression_ An expression that returns a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that will be printed.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]