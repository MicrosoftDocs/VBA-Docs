---
title: Application.ProjectBeforeTaskDelete2 event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeTaskDelete2
ms.assetid: 2c695579-bfe4-d109-eebc-4fb258a95c1e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeTaskDelete2 event (Project)

Occurs before a task is deleted. Uses the  **EventInfo** object parameter.


## Syntax

_expression_. `ProjectBeforeTaskDelete2`( `_tsk_`, `_Info_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _tsk_|Required|**Task**| The task that is being deleted.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the task is not deleted when the procedure is finished.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application. 

The  **ProjectBeforeTaskDelete2** event does not occur when changes have been made using a custom form.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]