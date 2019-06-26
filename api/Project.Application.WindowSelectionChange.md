---
title: Application.WindowSelectionChange event (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowSelectionChange
ms.assetid: 239c0a87-7966-b4b5-5731-9fe059f56a43
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowSelectionChange event (Project)

Occurs when the selection handle is changed within a window in Project.


## Syntax

_expression_. `WindowSelectionChange`( `_Window_`, `_sel_`, `_selType_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**|The window where the selection occurs.|
| _sel_|Required|**Selection**|The selection.|
| _selType_|Required|**Long**|The type of data included in the selection. Can be one of the following  **PjItemType** constants: **pjOtherItem**, **pjResourceItem**, or **pjTaskItem**.|

## Return value

**Nothing**


## Remarks

The  **WindowSelectionChange** event does not occur when changing the selection on the right pane of a **Task Usage** or **Resource Usage** view, or when changing the selection within a node in the **Network Diagram** view.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]