---
title: Application.ProjectBeforeResourceDelete event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeResourceDelete
ms.assetid: aadef12e-57dc-210e-d29a-54f79d1c1abd
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeResourceDelete event (Project)

Occurs before a resource is deleted.


## Syntax

_expression_. `ProjectBeforeResourceDelete`( `_res_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _res_|Required|**Resource**| The resource that is being deleted.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the resource is not deleted.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

The **ProjectBeforeResourceDelete** event doesn't occur when changes have been made using a custom form.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]