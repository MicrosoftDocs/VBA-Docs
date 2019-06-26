---
title: Application.WindowSidepaneDisplayChange event (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowSidepaneDisplayChange
ms.assetid: 8c4c22f4-4005-eff5-2964-880982634e78
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowSidepaneDisplayChange event (Project)

Occurs when the user shows or hides the Project Guide.


## Syntax

_expression_. `WindowSidepaneDisplayChange`( `_Window_`, `_Close_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**| The window where the **Project Guide** is shown or hidden.|
| _Close_|Required|**Boolean**|**False** if the user is closing the **Project Guide**.  **True** if the user is showing the **Project Guide**.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.


> [!NOTE] 
> The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of a custom Project Guide for new development.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]