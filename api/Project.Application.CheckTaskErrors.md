---
title: Application.CheckTaskErrors method (Project)
keywords: vbapj.chm2257
f1_keywords:
- vbapj.chm2257
ms.prod: project-server
api_name:
- Project.Application.CheckTaskErrors
ms.assetid: 7b361295-993a-13b2-b9bb-26f149e16e72
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CheckTaskErrors method (Project)

Checks the task to ensure that required custom fields are filled and that the calendars have the enterprise calendars type. If the TaskID parameter is  **null**, all tasks are checked. .


## Syntax

_expression_. `CheckTaskErrors`( `_TaskID_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TaskID_|Optional|**Variant**|TaskID for the task or  **null**.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]