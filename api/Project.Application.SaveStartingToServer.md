---
title: Application.SaveStartingToServer event (Project)
ms.prod: project-server
api_name:
- Project.Application.SaveStartingToServer
ms.assetid: e9d19b19-b916-a85d-486a-4a8676998b6c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SaveStartingToServer event (Project)

Occurs when Project Professional starts to save project changes to the Project Server queue. 


## Syntax

_expression_. `SaveStartingToServer`( `_bstrName_`, `_bstrprojGuid_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstrName_|Required|**String**|Name of the project.|
| _bstrprojGuid_|Required|**String**|GUID of the project.|

## Return value

**Nothing**


## Remarks

The **SaveStartingToServer** event cannot be cancelled.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]