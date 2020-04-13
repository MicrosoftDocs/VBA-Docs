---
title: Application.WindowGoalAreaChange event (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowGoalAreaChange
ms.assetid: 1ae33d11-f8aa-e1a2-b59d-9736ce4a6283
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowGoalAreaChange event (Project)

Occurs after a user clicks a different goal area in the Project Guide.


## Syntax

_expression_. `WindowGoalAreaChange`( `_Window_`, `_goalArea_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**|The window where the **Project Guide** is being changed.|
| _goalArea_|Required|**Long**|The ID of the goal area the user just clicked.|

## Return value

**Nothing**


## Remarks


> [!NOTE] 
> The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of a custom Project Guide for new development.

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]