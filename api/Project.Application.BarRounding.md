---
title: Application.BarRounding method (Project)
keywords: vbapj.chm2080
f1_keywords:
- vbapj.chm2080
ms.prod: project-server
api_name:
- Project.Application.BarRounding
ms.assetid: 6f776070-0a37-a72b-8cf8-ea3fd2c3fd06
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BarRounding method (Project)

Controls whether the start times of tasks are reflected by their corresponding task bars or the task bars are rounded to full days.


## Syntax

_expression_. `BarRounding`( `_On_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _On_|Optional|**Boolean**|**True** if bars round to the nearest day. The default value is **True**.|

## Return value

 **Boolean**


## Remarks

The **BarRounding** method affects only how tasks display on the Gantt Chart or Calendar. The duration of the tasks is not affected.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]