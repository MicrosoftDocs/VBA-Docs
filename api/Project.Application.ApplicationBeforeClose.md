---
title: Application.ApplicationBeforeClose event (Project)
ms.prod: project-server
api_name:
- Project.Application.ApplicationBeforeClose
ms.assetid: 9523a793-b4c1-fd79-303e-b167d7f80025
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ApplicationBeforeClose event (Project)

Occurs before Project exits.


## Syntax

_expression_. `ApplicationBeforeClose`( `_Info_`, )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Info_|Required|**EventInfo**|**EventInfo.Cancel** is **False** when the event occurs. If the event procedure sets this argument to **True**, Project does not close when the procedure is finished.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]