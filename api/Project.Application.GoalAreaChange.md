---
title: Application.GoalAreaChange method (Project)
keywords: vbapj.chm51
f1_keywords:
- vbapj.chm51
ms.prod: project-server
api_name:
- Project.Application.GoalAreaChange
ms.assetid: 84341db8-3f8e-44f3-4b34-e702ee2841dd
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GoalAreaChange method (Project)

Changes Project Guide goal areas by triggering the  **[WindowGoalAreaChange](Project.Application.WindowGoalAreaChange.md)** event. Deprecated in Project.


## Syntax

_expression_. `GoalAreaChange`( `_goalArea_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _goalArea_|Required|**Integer**|An integer that corresponds to the new goal area to which you are changing. For example, setting the  _goalArea_ argument to 1 will switch to the first goal area in the Project Guide.|

## Return value

 **Boolean**


## Remarks


> [!NOTE] 
> The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of the Project Guide for new development.

Script in the main.html page looks up and loads the appropriate task list page for the new goal area.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]