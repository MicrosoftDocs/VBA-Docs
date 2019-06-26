---
title: Application.WorkpaneDisplayChange event (Project)
ms.prod: project-server
api_name:
- Project.Application.WorkpaneDisplayChange
ms.assetid: 8fad51ed-57f5-a34d-6ef6-f699b605c10c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WorkpaneDisplayChange event (Project)

Occurs when the Project Guide is hidden or shown.


## Syntax

_expression_. `WorkpaneDisplayChange`( `_DisplayState_`, )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DisplayState_|Required|**Boolean**|**True** if the **Project Guide** is shown. **False** if the **Project Guide** is hidden.|

## Return value

**Nothing**


## Remarks


> [!NOTE] 
> The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of a custom Project Guide for new development.

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]