---
title: Application.ProjectBeforeTaskDelete event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeTaskDelete
ms.assetid: 3acc4ba4-0fdc-61fd-17df-e6450055a39b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeTaskDelete event (Project)

Occurs before a task is deleted.


## Syntax

_expression_. `ProjectBeforeTaskDelete`( `_tsk_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _tsk_|Required|**Task**| The task that is being deleted.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the task is not deleted.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application. The  **ProjectBeforeTaskDelete** event does not occur when changes have been made using a custom form.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]