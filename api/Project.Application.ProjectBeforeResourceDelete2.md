---
title: Application.ProjectBeforeResourceDelete2 event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeResourceDelete2
ms.assetid: 3665f6e0-6df8-0a8d-28c1-49bfe51ffad5
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeResourceDelete2 event (Project)

Occurs before a resource is deleted. Uses the  **EventInfo** object parameter.


## Syntax

_expression_. `ProjectBeforeResourceDelete2`( `_res_`, `_Info_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _res_|Required|**Resource**| The resource that is being deleted.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the resource is not deleted.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

The  **ProjectBeforeResourceDelete2** event doesn't occur when changes have been made using a custom form.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]