---
title: Application.DDEInitiate method (Project)
keywords: vbapj.chm1201
f1_keywords:
- vbapj.chm1201
ms.prod: project-server
api_name:
- Project.Application.DDEInitiate
ms.assetid: a517c66f-4bec-9bec-270c-2053bc733145
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DDEInitiate method (Project)

Opens a dynamic data exchange (DDE) channel to an application.


## Syntax

_expression_. `DDEInitiate`( `_App_`, `_Topic_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _App_|Required|**String**| The name of the application to which you want to send commands.|
| _Topic_|Required|**String**|A document in the application to which you want to send commands.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]