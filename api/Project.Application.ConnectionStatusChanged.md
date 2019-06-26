---
title: Application.ConnectionStatusChanged event (Project)
ms.prod: project-server
api_name:
- Project.Application.ConnectionStatusChanged
ms.assetid: ffc6fc8a-f5b7-3a3d-4829-712a8305ed17
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ConnectionStatusChanged event (Project)

Occurs when the status of the connection with Project Server changes. Available only in Project Professional.


## Syntax

_expression_. `ConnectionStatusChanged`( `_online_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _online_|Required|**Boolean**|**True** if Project Professional is connected with Project Server; otherwise, **False**.|

## Return value

**Nothing**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]