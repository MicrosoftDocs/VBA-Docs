---
title: Application.ProjectBeforeResourceNew event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeResourceNew
ms.assetid: a432c713-d1fa-0743-ff4e-90fbd724dfe4
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeResourceNew event (Project)

Occurs before one or more resources are created.


## Syntax

_expression_. `ProjectBeforeResourceNew`( `_pj_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project in which the resource or resources are being created.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the new resource or resources are not created.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

The  **ProjectBeforeResourceNew** event doesn't occur during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]