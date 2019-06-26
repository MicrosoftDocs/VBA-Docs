---
title: Application.GoalAreaTaskHighlight method (Project)
ms.prod: project-server
api_name:
- Project.Application.GoalAreaTaskHighlight
ms.assetid: 32616617-d34a-c9f4-8ddd-17fa3f1c7e74
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GoalAreaTaskHighlight method (Project)

Highlights a specified task in the Project Guide. Deprecated in Project.


## Syntax

_expression_. `GoalAreaTaskHighlight`( `_TaskID_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TaskID_|Required|**Long**|The Task ID you wish to highlight.|

## Remarks


> [!NOTE] 
> The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of the Project Guide for new development.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]