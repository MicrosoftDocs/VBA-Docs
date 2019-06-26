---
title: Application.SaveCompletedToServer event (Project)
ms.prod: project-server
api_name:
- Project.Application.SaveCompletedToServer
ms.assetid: 05ca27a0-a6cd-efbd-eff8-4f457c3de5c0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SaveCompletedToServer event (Project)

Occurs when Project Professional successfully puts the  **Project Save** job in the Project Server Queue.


## Syntax

_expression_. `SaveCompletedToServer`( `_bstrName_`, `_bstrprojGuid_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstrName_|Required|**String**|Name of the project.|
| _bstrprojGuid_|Required|**String**|GUID of the project|

## Return value

**Nothing**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]