---
title: Application.WindowActivate event (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowActivate
ms.assetid: b54d0956-7eab-db5f-394a-5120bc111afd
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowActivate event (Project)

Occurs when any window within Project is activated. The  **WindowActivate** event does not occur when the application window is activated.


## Syntax

_expression_. `WindowActivate`( `_activatedWindow_`, )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _activatedWindow_|Required|**Window**|The activated window.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]