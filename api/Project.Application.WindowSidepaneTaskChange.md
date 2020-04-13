---
title: Application.WindowSidepaneTaskChange event (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowSidepaneTaskChange
ms.assetid: 674a8134-1e34-2658-6c67-5eb92c628ed8
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowSidepaneTaskChange event (Project)

Occurs when a user selects different items in the **Next Steps and Related Activities** menu in the Project Guide.


## Syntax

_expression_. `WindowSidepaneTaskChange`( `_Window_`, `_ID_`, `_IsGoalArea_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**|The window where the **Project Guide** is being changed.|
| _ID_|Required|**Long**|The ID of the task in the **Project Guide** you are trying to display.|
| _IsGoalArea_|Required|**Boolean**|**True** if trying to change to a different goal area in the **Project Guide**.  **False** if trying to change to a different **Project Guide** task.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.


> [!NOTE] 
> The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of a custom Project Guide for new development.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]