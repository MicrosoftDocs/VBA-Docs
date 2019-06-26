---
title: Application.ProjectBeforeSave event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeSave
ms.assetid: 406986e7-22f6-109e-1973-f22e81081111
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeSave event (Project)

Occurs before a project is saved.


## Syntax

_expression_. `ProjectBeforeSave`( `_pj_`, `_SaveAsUi_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project to be saved.|
| _SaveAsUi_|Required|**Boolean**|**True** if the **Save As** dialog box is displayed.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the project will not be saved.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]