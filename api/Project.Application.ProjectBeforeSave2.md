---
title: Application.ProjectBeforeSave2 event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeSave2
ms.assetid: 5afcdb4c-85e6-183c-f6e7-333d2a7ea3d4
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeSave2 event (Project)

Occurs before a project is saved. Uses the **EventInfo** object parameter.


## Syntax

_expression_. `ProjectBeforeSave2`( `_pj_`, `_SaveAsUi_`, `_Info_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project to be saved.|
| _SaveAsUi_|Required|**Boolean**|**True** if the **Save As** dialog box is displayed.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is **False** when the event occurs. If the event procedure sets this argument to **True**, the project will not be saved when the procedure is finished.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]