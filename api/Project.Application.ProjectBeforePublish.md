---
title: Application.ProjectBeforePublish event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforePublish
ms.assetid: 5778ec6c-a8c0-0a05-145c-c9ad6132bf87
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforePublish event (Project)

Occurs before a  **Publish** operation is placed on the server queue. The **ProjectBeforePublish** event can be cancelled. Project Professional only.


## Syntax

_expression_. `ProjectBeforePublish`( `_pj_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|Project object.|
| _Cancel_|Required|**Boolean**|**True** to cancel the **Publish** job.|

## Return value

Nothing


## Remarks

The  **ProjectBeforePublish** event is commonly used to determine whether certain conditions are satisfied and to cancel publishing if the conditions are not met.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]