---
title: Project.Deactivate event (Project)
keywords: vbapj.chm131189
f1_keywords:
- vbapj.chm131189
ms.prod: project-server
api_name:
- Project.Project.Deactivate
ms.assetid: ce4301e5-8881-1280-fafb-a87c37d088dd
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Deactivate event (Project)

Occurs when switching from the current project to another project.


## Syntax

_expression_. `Deactivate`( `_pj_` )

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that was deactivated.|

## Return value

**Nothing**


## Remarks

The **Deactivate** event does not occur when you close a project or when you switch between two windows showing the same project.

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]